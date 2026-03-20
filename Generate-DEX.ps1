#Requires -Version 5.1
<#
.SYNOPSIS
    Generate-DEX.ps1 - Génération automatique du Document d'Exploitation (DEX) système
    au format HTML pour les serveurs Windows Server 2016 et supérieur.

.DESCRIPTION
    Ce script collecte les informations système d'un serveur Windows et génère un rapport
    HTML structuré (DEX - Document d'EXploitation). Il analyse également les EventLogs
    pour déduire les plages de reboot programmées. Un emplacement optionnel peut être
    fourni pour inclure le planning de sauvegarde.

.PARAMETER OutputPath
    Chemin du dossier de sortie pour le fichier HTML. Par défaut : C:\Temp\DEX

.PARAMETER Redactor
    Nom du rédacteur du document. Par défaut : $env:USERNAME

.PARAMETER RedactorTitle
    Fonction/titre du rédacteur. Par défaut : "System Engineer"

.PARAMETER BackupScheduleFile
    Chemin vers un fichier texte ou CSV décrivant le planning de sauvegarde (optionnel).

.PARAMETER DaysToAnalyzeReboots
    Nombre de jours d'historique EventLog à analyser pour détecter les reboots planifiés.
    Par défaut : 90

.EXAMPLE
    .\Generate-DEX.ps1
    .\Generate-DEX.ps1 -OutputPath "D:\Reports\DEX" -Redactor "Hicham Mouzoun" -DaysToAnalyzeReboots 60
    .\Generate-DEX.ps1 -BackupScheduleFile "C:\Temp\backup-schedule.csv"

.NOTES
    Auteur  : Ingénierie Système
    Version : 1.0
    Requis  : Windows Server 2016+, PowerShell 5.1+, droits administrateur locaux
#>

[CmdletBinding()]
param(
    [string]$OutputPath       = "C:\Temp\DEX",
    [string]$Redactor         = $env:USERNAME,
    [string]$RedactorTitle    = "System Engineer",
    [string]$BackupScheduleFile = "",
    [int]   $DaysToAnalyzeReboots = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

#region ── Helpers ─────────────────────────────────────────────────────────────

function Escape-Html {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Html-Row {
    param([string]$Key, [string]$Value, [bool]$Highlight = $false)
    $cls = if ($Highlight) { " class='highlight'" } else { "" }
    "<tr$cls><th>$(Escape-Html $Key)</th><td>$(Escape-Html $Value)</td></tr>"
}

function Html-Section {
    param([string]$Title)
    "<h2>$(Escape-Html $Title)</h2>"
}

#endregion

#region ── Collecte des données système ────────────────────────────────────────

Write-Host "[*] Collecte des informations système..." -ForegroundColor Cyan

# --- Infos OS & Machine ---
$cs   = Get-CimInstance Win32_ComputerSystem
$os   = Get-CimInstance Win32_OperatingSystem
$bios = Get-CimInstance Win32_BIOS

$serverName = $env:COMPUTERNAME
$osCaption  = $os.Caption
$osVersion  = $os.Version
$domain     = $cs.Domain
$mfr        = $cs.Manufacturer
$model      = $cs.Model

# --- Processeurs ---
$cpuList = Get-CimInstance Win32_Processor
$totalSockets = ($cpuList | Measure-Object).Count
$totalCores   = ($cpuList | Measure-Object -Property NumberOfCores -Sum).Sum
$totalThreads = ($cpuList | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum

# Détection hyperviseur
$hypervisor = ""
if ($model -match "VMware")       { $hypervisor = "VMware ESXi" }
elseif ($model -match "Virtual")  { $hypervisor = "Microsoft Hyper-V" }
elseif ($bios.SerialNumber -match "Xen") { $hypervisor = "Xen" }
else                               { $hypervisor = "Physique / Inconnu" }

$cpuSectionLabel = if ($hypervisor -ne "Physique / Inconnu") { "Processeur ($hypervisor)" } else { "Processeur" }

# --- RAM ---
$ramBanks = Get-CimInstance Win32_PhysicalMemory

# --- Disques physiques ---
$diskPhys = Get-CimInstance Win32_DiskDrive

# --- Volumes logiques ---
$volumes = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -in @(2,3) }
# Inclure les partitions sans lettre (System Reserved, etc.)
$partitions = Get-CimInstance Win32_Volume | Where-Object { $_.DriveLetter -eq $null -and $_.Label -ne "" }

# --- Cartes réseau ---
$nics = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }

# --- Variables d'environnement ---
$envVars = [System.Environment]::GetEnvironmentVariables([System.EnvironmentVariableTarget]::Machine)
$envVarsUser = [System.Environment]::GetEnvironmentVariables([System.EnvironmentVariableTarget]::Process)
# On fusionne et trie
$allEnvKeys = ($envVars.Keys + $envVarsUser.Keys) | Select-Object -Unique | Sort-Object

# --- Stratégie mot de passe ---
$netAccounts = net accounts 2>&1

# --- Membres du groupe Administrators ---
try {
    $admins = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue
} catch {
    $admins = @()
}

# --- Applications installées ---
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$installedApps = foreach ($p in $regPaths) {
    Get-ItemProperty $p -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -and $_.DisplayName -ne "" } |
        Select-Object DisplayName, DisplayVersion
}
$installedApps = $installedApps | Sort-Object DisplayName -Unique

# Apps à mettre en évidence (versions obsolètes/connues sensibles)
$highlightPatterns = @(
    "Microsoft Visual C\+\+ 200[0-8]",
    "Microsoft \.NET.* 5\.0",
    "Python 3\.[0-8]\.",
    "Node\.js\s+1[0-6]\.",
    "Java.*1\.[0-8]\.",
    "OpenSSL 1\."
)

# --- Tâches planifiées ---
$scheduledTasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
    Where-Object {
        $_.TaskPath -notmatch "^\\Microsoft\\" -and
        $_.TaskPath -notmatch "^\\Windows\\"
    } |
    Select-Object TaskName, TaskPath,
        @{N="Author"; E={ $_.Principal.UserId }},
        @{N="Status"; E={ $_.State }}

#endregion

#region ── Analyse EventLog : détection des reboots planifiés ──────────────────

Write-Host "[*] Analyse des EventLogs pour détecter les reboots planifiés..." -ForegroundColor Cyan

$rebootScheduleHtml = ""

try {
    $startDate = (Get-Date).AddDays(-$DaysToAnalyzeReboots)

    # EventID 6005 = démarrage du service EventLog (= fin de reboot)
    # EventID 6006 = arrêt propre (shutdown)
    # EventID 1074 = reboot/shutdown initié par un utilisateur ou process
    $rebootEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Id        = @(6005, 6006, 1074)
        StartTime = $startDate
    } -ErrorAction SilentlyContinue

    if ($rebootEvents -and $rebootEvents.Count -gt 0) {

        # Regrouper par heure (tranche horaire) pour identifier les créneaux répétitifs
        $bootEvents = $rebootEvents | Where-Object { $_.Id -eq 6005 }

        $hourGroups = $bootEvents |
            Group-Object { $_.TimeCreated.Hour } |
            Sort-Object Count -Descending

        $dayGroups = $bootEvents |
            Group-Object { $_.TimeCreated.DayOfWeek } |
            Sort-Object Count -Descending

        # Construire le tableau de synthèse
        $totalReboots = $bootEvents.Count

        $rebootTableRows = ""
        foreach ($hg in $hourGroups | Select-Object -First 5) {
            $pct = [math]::Round(($hg.Count / $totalReboots) * 100, 1)
            $bar = "█" * [math]::Round($pct / 5)
            $rebootTableRows += "<tr><td>{0}h00 - {1}h00</td><td>$($hg.Count)</td><td>$pct %</td><td style='font-family:monospace;color:#0074A2'>$bar</td></tr>" -f $hg.Name, ([int]$hg.Name + 1)
        }

        $dayTableRows = ""
        foreach ($dg in $dayGroups | Select-Object -First 7) {
            $pct = [math]::Round(($dg.Count / $totalReboots) * 100, 1)
            $dayTableRows += "<tr><td>$($dg.Name)</td><td>$($dg.Count)</td><td>$pct %</td></tr>"
        }

        # Détecter la plage la plus probable
        $topHour = $hourGroups | Select-Object -First 1
        $topDay  = $dayGroups  | Select-Object -First 1
        $likelyWindow = ""
        if ($topHour) {
            $likelyWindow = "Créneau le plus probable : <strong>{0}h00 – {1}h00</strong>" -f $topHour.Name, ([int]$topHour.Name + 1)
            if ($topDay) { $likelyWindow += " le <strong>$($topDay.Name)</strong>" }
            $likelyWindow += " (sur $totalReboots reboots analysés sur $DaysToAnalyzeReboots jours)"
        }

        # Derniers reboots
        $lastRebootRows = ""
        $rebootEvents | Where-Object { $_.Id -eq 1074 } | Select-Object -First 10 | ForEach-Object {
            $msg = $_.Message -replace '\s+',' '
            $lastRebootRows += "<tr><td>$(Escape-Html $_.TimeCreated.ToString('dd/MM/yyyy HH:mm:ss'))</td><td>$(Escape-Html ($msg.Substring(0,[math]::Min(120,$msg.Length))))</td></tr>"
        }
        if (-not $lastRebootRows) {
            $bootEvents | Select-Object -First 10 | ForEach-Object {
                $lastRebootRows += "<tr><td>$(Escape-Html $_.TimeCreated.ToString('dd/MM/yyyy HH:mm:ss'))</td><td>Démarrage du service EventLog (EventID 6005)</td></tr>"
            }
        }

        $rebootScheduleHtml = @"
<h2>Analyse des Reboots (EventLog - $DaysToAnalyzeReboots derniers jours)</h2>
<p style='color:#333'>$likelyWindow</p>
<table>
  <tr><th colspan='4' style='background:#e8f4fc'>Distribution horaire des reboots</th></tr>
  <tr><th>Créneau horaire</th><th>Occurrences</th><th>Fréquence</th><th>Distribution</th></tr>
  $rebootTableRows
</table>
<table>
  <tr><th colspan='3' style='background:#e8f4fc'>Distribution par jour de la semaine</th></tr>
  <tr><th>Jour</th><th>Occurrences</th><th>Fréquence</th></tr>
  $dayTableRows
</table>
<table>
  <tr><th colspan='2' style='background:#e8f4fc'>Derniers événements de reboot (EventID 1074 / 6005)</th></tr>
  <tr><th>Date/Heure</th><th>Détail</th></tr>
  $lastRebootRows
</table>
"@
    } else {
        $rebootScheduleHtml = "<h2>Analyse des Reboots</h2><p><em>Aucun événement de reboot trouvé dans les $DaysToAnalyzeReboots derniers jours.</em></p>"
    }
} catch {
    $rebootScheduleHtml = "<h2>Analyse des Reboots</h2><p><em>Impossible de lire les EventLogs : $($_.Exception.Message)</em></p>"
}

#endregion

#region ── Backup Schedule (optionnel) ─────────────────────────────────────────

$backupScheduleHtml = ""

if ($BackupScheduleFile -and (Test-Path $BackupScheduleFile)) {
    Write-Host "[*] Chargement du planning de sauvegarde depuis : $BackupScheduleFile" -ForegroundColor Cyan
    $ext = [System.IO.Path]::GetExtension($BackupScheduleFile).ToLower()

    if ($ext -eq ".csv") {
        try {
            $bkpData = Import-Csv -Path $BackupScheduleFile -Delimiter ";" -ErrorAction Stop
            $headers = $bkpData[0].PSObject.Properties.Name
            $bkpRows = ""
            foreach ($row in $bkpData) {
                $bkpRows += "<tr>"
                foreach ($h in $headers) { $bkpRows += "<td>$(Escape-Html $row.$h)</td>" }
                $bkpRows += "</tr>"
            }
            $bkpHeaders = ($headers | ForEach-Object { "<th>$(Escape-Html $_)</th>" }) -join ""
            $backupScheduleHtml = @"
<h2>Planning de Sauvegarde</h2>
<table>
  <tr>$bkpHeaders</tr>
  $bkpRows
</table>
"@
        } catch {
            $backupScheduleHtml = "<h2>Planning de Sauvegarde</h2><p><em>Erreur de lecture CSV : $($_.Exception.Message)</em></p>"
        }
    } else {
        # Fichier texte brut : affichage dans un bloc <pre>
        $bkpContent = Get-Content $BackupScheduleFile -Raw -ErrorAction SilentlyContinue
        $backupScheduleHtml = @"
<h2>Planning de Sauvegarde</h2>
<pre style='background:#f0f0f0;padding:10px;border:1px solid #ccc;overflow:auto'>$(Escape-Html $bkpContent)</pre>
"@
    }
} else {
    # Section vide prête à être complétée
    $backupScheduleHtml = @"
<h2>Planning de Sauvegarde</h2>
<table>
  <tr><th>Job</th><th>Type</th><th>Fréquence</th><th>Heure</th><th>Rétention</th><th>Cible</th><th>Statut</th></tr>
  <tr><td colspan='7' style='text-align:center;color:#999'><em>À compléter — utilisez -BackupScheduleFile pour charger automatiquement depuis un CSV</em></td></tr>
</table>
"@
}

#endregion

#region ── Construction des sections HTML ──────────────────────────────────────

Write-Host "[*] Génération du contenu HTML..." -ForegroundColor Cyan

$now          = Get-Date
$dateStr      = $now.ToString("dd/MM/yyyy HH:mm:ss")
$fileStamp    = $now.ToString("yyyyMMdd-HHmm")
$fileName     = "DEX-Report-$fileStamp.html"
$fullPath     = Join-Path $OutputPath $fileName

# ── Informations Système ──────────────────────────────────────────────────────
$sysInfoHtml = "<h2>Informations Système</h2><table>"
$sysInfoHtml += Html-Row "Nom"        $serverName
$sysInfoHtml += Html-Row "Système"    $osCaption
$sysInfoHtml += Html-Row "Version"    $osVersion
$sysInfoHtml += Html-Row "Fabricant"  $mfr
$sysInfoHtml += Html-Row "Modèle"     $model
$sysInfoHtml += Html-Row "Domaine"    $domain
$sysInfoHtml += "</table>"

# ── Processeurs ───────────────────────────────────────────────────────────────
$cpuHtml = "<h2>$cpuSectionLabel</h2><table>"
foreach ($cpu in $cpuList) {
    $cpuHtml += Html-Row "Virtual CPU (vCPU)"  $cpu.Name
    $cpuHtml += Html-Row "vCores"              $cpu.NumberOfCores
    $cpuHtml += Html-Row "vThreads"            $cpu.NumberOfLogicalProcessors
}
$cpuHtml += Html-Row "Total vCPU (sockets)" $totalSockets
$cpuHtml += Html-Row "Total vCores"         $totalCores
$cpuHtml += Html-Row "Total vThreads"       $totalThreads
$cpuHtml += Html-Row "Hyperviseur"          $hypervisor
$cpuHtml += "</table>"

# ── Mémoire RAM ───────────────────────────────────────────────────────────────
$ramHtml = "<h2>Mémoire RAM</h2><table><tr><th>Capacité</th><th>Type</th><th>Fréquence</th></tr>"
if ($ramBanks) {
    foreach ($bank in $ramBanks) {
        $cap   = Format-Bytes $bank.Capacity
        $type  = switch ($bank.MemoryType) {
            20 { "DDR" } 21 { "DDR2" } 22 { "DDR2 FB-DIMM" } 24 { "DDR3" } 26 { "DDR4" } 34 { "DDR5" }
            default { "Inconnu ($($bank.MemoryType))" }
        }
        $freq  = if ($bank.Speed) { "$($bank.Speed) MHz" } else { "- MHz" }
        $ramHtml += "<tr><td>$cap</td><td>$(Escape-Html $type)</td><td>$(Escape-Html $freq)</td></tr>"
    }
} else {
    $totalRAM = Format-Bytes ($cs.TotalPhysicalMemory)
    $ramHtml += "<tr><td>$totalRAM</td><td>Virtuelle (VM)</td><td>-</td></tr>"
}
$ramHtml += "</table>"

# ── Disques Physiques ─────────────────────────────────────────────────────────
$diskHtml = "<h2>Disques Physiques</h2><table><tr><th>Modèle</th><th>Capacité</th><th>Type</th></tr>"
foreach ($d in $diskPhys) {
    $size = Format-Bytes $d.Size
    $diskHtml += "<tr><td>$(Escape-Html $d.Model)</td><td>$size</td><td>$(Escape-Html $d.MediaType)</td></tr>"
}
$diskHtml += "</table>"

# ── Volumes Logiques ──────────────────────────────────────────────────────────
$volHtml = "<h2>Volumes Logiques</h2><table><tr><th>Lettre</th><th>Label</th><th>Système de Fichiers</th><th>Taille</th><th>Espace Libre</th></tr>"
foreach ($v in $volumes) {
    $letter = $v.DeviceID -replace ":\\",""
    $size   = Format-Bytes $v.Size
    $free   = Format-Bytes $v.FreeSpace
    $volHtml += "<tr><td>$(Escape-Html $letter)</td><td>$(Escape-Html $v.VolumeName)</td><td>$(Escape-Html $v.FileSystem)</td><td>$size</td><td>$free</td></tr>"
}
foreach ($p in $partitions) {
    $size = Format-Bytes $p.Capacity
    $free = Format-Bytes $p.FreeSpace
    $volHtml += "<tr><td></td><td>$(Escape-Html $p.Label)</td><td>$(Escape-Html $p.FileSystem)</td><td>$size</td><td>$free</td></tr>"
}
$volHtml += "</table>"

# ── Cartes Réseau ─────────────────────────────────────────────────────────────
$nicHtml = "<h2>Cartes Réseau</h2><table><tr><th>Nom</th><th>Adresse MAC</th><th>Adresse IP</th></tr>"
foreach ($nic in $nics) {
    $ips  = ($nic.IPAddress | Where-Object { $_ -match '\d+\.\d+\.\d+\.\d+' }) -join ", "
    $desc = $nic.Description
    $mac  = $nic.MACAddress
    $nicHtml += "<tr><td>$(Escape-Html $desc)</td><td>$(Escape-Html $mac)</td><td>$(Escape-Html $ips)</td></tr>"
}
$nicHtml += "</table>"

# ── Variables d'environnement ─────────────────────────────────────────────────
$envHtml = "<h2>Variables d'environnement</h2><table><tr><th>Nom</th><th>Valeur</th></tr>"
foreach ($key in $allEnvKeys) {
    $val = if ($envVars[$key]) { $envVars[$key] } else { $envVarsUser[$key] }
    $envHtml += "<tr><td>$(Escape-Html $key)</td><td>$(Escape-Html $val)</td></tr>"
}
$envHtml += "</table>"

# ── Stratégie mot de passe ────────────────────────────────────────────────────
$pwdHtml = "<h2>Stratégie de mot de passe locale</h2><table><tr><th>Paramètre</th><th>Valeur</th></tr>"
foreach ($line in $netAccounts) {
    if ($line -match "^(.+?)\s{2,}(.+)$") {
        $pwdHtml += "<tr><td>$(Escape-Html $Matches[1].Trim())</td><td>$(Escape-Html $Matches[2].Trim())</td></tr>"
    }
}
$pwdHtml += "</table>"

# ── Membres du groupe Administrators ─────────────────────────────────────────
$admHtml = "<h2>Membres du groupe Administrators</h2><table><tr><th>Nom</th><th>Type</th></tr>"
foreach ($adm in $admins) {
    $admHtml += "<tr><td>$(Escape-Html $adm.Name)</td><td>$(Escape-Html $adm.ObjectClass)</td></tr>"
}
$admHtml += "</table>"

# ── Applications Installées ───────────────────────────────────────────────────
$appHtml = "<h2>Applications Installées</h2><table><tr><th>Nom</th><th>Version</th></tr>"
foreach ($app in $installedApps) {
    $isHighlight = $false
    foreach ($pat in $highlightPatterns) {
        if ($app.DisplayName -match $pat) { $isHighlight = $true; break }
    }
    $cls = if ($isHighlight) { " class='highlight'" } else { "" }
    $appHtml += "<tr$cls><td>$(Escape-Html $app.DisplayName)</td><td>$(Escape-Html $app.DisplayVersion)</td></tr>"
}
$appHtml += "</table>"

# ── Tâches Planifiées ─────────────────────────────────────────────────────────
$taskHtml = "<h2>Tâches Planifiées Personnalisées</h2><table><tr><th>Nom</th><th>Chemin</th><th>Auteur</th><th>Statut</th></tr>"
foreach ($t in $scheduledTasks) {
    $taskHtml += "<tr><td>$(Escape-Html $t.TaskName)</td><td>$(Escape-Html $t.TaskPath)</td><td>$(Escape-Html $t.Author)</td><td>$(Escape-Html $t.Status)</td></tr>"
}
$taskHtml += "</table>"

#endregion

#region ── Template HTML complet ───────────────────────────────────────────────

$html = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Language" content="fr">
  <title>DEX $serverName</title>
  <style>
    body { font-family: Arial; font-size: 13px; margin: 30px; background-color: #f9f9f9; }
    .container { border: 2px solid #ccc; border-radius: 8px; padding: 20px;
                 background-color: #fff; box-shadow: 2px 2px 8px #999; }
    h1 { color: #003366; text-align: center; }
    h2 { border-bottom: 1px solid #ccc; color: #489CDF; margin-top: 30px; }
    table { border-collapse: collapse; width: 80%; margin-bottom: 20px; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
    th { background-color: #e8f4fc; }
    .highlight { background-color: #ffffcc; font-weight: bold; }
    pre { white-space: pre-wrap; word-wrap: break-word; }
    .cover-table { text-align: center; width: 50%; }
    .cover-table td { text-align: left; }
    .priority-table { border-collapse: collapse; text-align: center; }
    .priority-table th { background-color: #0074A2; color: white; padding: 6px; border: 1px solid #ccc; }
    .priority-table td { padding: 6px; border: 1px solid #ccc; }
    .priority-table .header-cell { background-color: #0074A2; color: black; }
    @media print {
      .container { box-shadow: none; border: 1px solid #999; }
    }
  </style>
</head>
<body>
<div class="container">

  <!-- ═══════════════ PAGE DE GARDE ═══════════════ -->
  <br><br><br><br><br><br><br><br><br><br>
  <h1>Engineering Systems Documentation</h1>
  <h1>$serverName</h1>

  <center style="margin-top:150px">
    <br><br>
    <table class="cover-table">
      <tbody>
        <tr>
          <td><span style="font-weight:bold">Version :</span></td>
          <td>1.0</td>
        </tr>
        <tr>
          <td><span style="font-weight:bold">Document date :</span></td>
          <td>$dateStr</td>
        </tr>
        <tr>
          <td><span style="font-weight:bold">Document filename :</span></td>
          <td>$fullPath</td>
        </tr>
      </tbody>
    </table>
  </center>

  <br><br><br><br>

  <!-- ═══════════════ APPROBATION ═══════════════ -->
  <h2>Informations d'approbation</h2>
  <table>
    <tbody>
      <tr>
        <td></td>
        <td><span style="font-weight:bold">Name</span></td>
        <td><span style="font-weight:bold">Function</span></td>
      </tr>
      <tr>
        <td>Redactor :</td>
        <td>$(Escape-Html $Redactor)</td>
        <td>$(Escape-Html $RedactorTitle)</td>
      </tr>
      <tr><td>Verified by :</td><td></td><td></td></tr>
      <tr>
        <td>Approver :<br>Visa :<br>Date :</td>
        <td></td><td></td>
      </tr>
    </tbody>
  </table>

  <!-- ═══════════════ GRILLE INCIDENTS ═══════════════ -->
  <h2>Grille de priorisation des incidents - BaseX -</h2>
  <table class="priority-table">
    <thead>
      <tr>
        <th rowspan="2">Impact</th>
        <th colspan="2">Urgence 1</th>
        <th colspan="2">Urgence 2</th>
        <th colspan="2">Urgence 3</th>
      </tr>
      <tr>
        <th>Urgence</th><th>Priorité</th>
        <th>Urgence</th><th>Priorité</th>
        <th>Urgence</th><th>Priorité</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="header-cell">Corp</td>
        <td>1</td><td>1</td><td>2</td><td>2</td><td>3</td><td>2</td>
      </tr>
      <tr>
        <td class="header-cell">Country/Region</td>
        <td>1</td><td>1</td><td>2</td><td>2</td><td>3</td><td>3</td>
      </tr>
      <tr>
        <td class="header-cell">User/Shop</td>
        <td>1</td><td>2</td><td>2</td><td>3</td><td>3</td><td>4</td>
      </tr>
    </tbody>
    <tfoot>
      <tr>
        <td class="header-cell">Criticité</td>
        <td colspan="2">C1</td><td colspan="2">C2</td><td colspan="2">C3</td>
      </tr>
    </tfoot>
  </table>
  <p>Target windows support : MA.WINDOWSServer&nbsp;&nbsp;&nbsp;
     criticality * Incident Severity = Incident Priority</p>

  <!-- ═══════════════ SECTIONS TECHNIQUES ═══════════════ -->
  $sysInfoHtml
  $cpuHtml
  $ramHtml
  $diskHtml
  $volHtml
  $nicHtml
  $envHtml
  $pwdHtml
  $admHtml
  $appHtml
  $taskHtml

  <!-- ═══════════════ REBOOTS PLANIFIÉS ═══════════════ -->
  $rebootScheduleHtml

  <!-- ═══════════════ PLANNING SAUVEGARDE ═══════════════ -->
  $backupScheduleHtml

</div>
</body>
</html>
"@

#endregion

#region ── Écriture du fichier ─────────────────────────────────────────────────

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Host "[+] Répertoire créé : $OutputPath" -ForegroundColor Green
}

$html | Out-File -FilePath $fullPath -Encoding UTF8 -Force

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  DEX généré avec succès !" -ForegroundColor Green
Write-Host "  Fichier : $fullPath" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green

# Ouvrir automatiquement dans le navigateur par défaut
try {
    Start-Process $fullPath -ErrorAction SilentlyContinue
} catch { }

#endregion
