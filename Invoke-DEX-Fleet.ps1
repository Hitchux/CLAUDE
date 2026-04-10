#Requires -Version 5.1
<#
.SYNOPSIS
    Invoke-DEX-Fleet.ps1 - Exécution du DEX sur une liste de serveurs distants.

.DESCRIPTION
    Lit une liste de serveurs depuis un fichier texte (un nom/IP par ligne),
    copie et exécute Generate-DEX.ps1 sur chaque machine via PowerShell Remoting
    (WinRM), puis rapatrie le fichier HTML généré en local.

    Prérequis :
      - WinRM activé sur les VMs cibles (Enable-PSRemoting -Force)
      - Droits administrateur locaux sur chaque VM
      - Le port 5985 (HTTP) ou 5986 (HTTPS) ouvert entre cette machine et les cibles

.PARAMETER VMList
    Chemin vers le fichier texte contenant la liste des serveurs (un par ligne).
    Lignes vides et commentaires (#) ignorés.
    Par défaut : .\vm.txt

.PARAMETER OutputPath
    Dossier local où seront copiés les fichiers HTML générés.
    Par défaut : .\DEX-Reports

.PARAMETER ScriptPath
    Chemin vers Generate-DEX.ps1. Par défaut : .\Generate-DEX.ps1

.PARAMETER Credential
    PSCredential à utiliser pour la connexion distante.
    Si omis, utilise les credentials courants (Kerberos/Pass-through).

.PARAMETER Redactor
    Nom du rédacteur inscrit dans chaque DEX. Par défaut : $env:USERNAME

.PARAMETER RedactorTitle
    Titre/fonction du rédacteur. Par défaut : "System Engineer"

.PARAMETER DaysToAnalyzeReboots
    Nombre de jours d'EventLog à analyser pour la détection de reboots.
    Par défaut : 90

.PARAMETER ThrottleLimit
    Nombre de sessions PS parallèles simultanées. Par défaut : 5

.PARAMETER UseSSL
    Utilise WinRM over HTTPS (port 5986) au lieu de HTTP (port 5985).

.PARAMETER SkipCACheck
    Ignore la vérification du certificat SSL (utile en lab / auto-signé).

.EXAMPLE
    # Avec credentials courants (domaine)
    .\Invoke-DEX-Fleet.ps1

    # Avec fichier VM personnalisé et credentials explicites
    $cred = Get-Credential
    .\Invoke-DEX-Fleet.ps1 -VMList "C:\Lists\servers.txt" -Credential $cred -ThrottleLimit 10

    # HTTPS avec certificat auto-signé
    .\Invoke-DEX-Fleet.ps1 -UseSSL -SkipCACheck -Credential (Get-Credential)

.NOTES
    Auteur  : Ingénierie Système
    Version : 1.0
#>

[CmdletBinding()]
param(
    [string]      $VMList               = ".\vm.txt",
    [string]      $OutputPath           = ".\DEX-Reports",
    [string]      $ScriptPath           = ".\Generate-DEX.ps1",
    [PSCredential]$Credential           = $null,
    [string]      $Redactor             = $env:USERNAME,
    [string]      $RedactorTitle        = "System Engineer",
    [int]         $DaysToAnalyzeReboots = 90,
    [int]         $ThrottleLimit        = 5,
    [switch]      $UseSSL,
    [switch]      $SkipCACheck,

    # ── Email via Power Automate Flow (contournement DLP) ─────────────────────
    # Prérequis : créer un Flow avec déclencheur HTTP + action "Envoyer un email (V2)"
    # Voir Send-DEXReport.ps1 pour le schéma JSON attendu par le déclencheur.
    [switch]      $SendEmail,
    [string]      $FlowWebhookUrl      = "",
    [string[]]    $EmailTo             = @(),
    [string]      $EmailCc             = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

#region ── Validation des entrées ──────────────────────────────────────────────

if (-not (Test-Path $VMList)) {
    Write-Error "Fichier VM introuvable : $VMList"
    exit 1
}

if (-not (Test-Path $ScriptPath)) {
    Write-Error "Script Generate-DEX.ps1 introuvable : $ScriptPath"
    exit 1
}

$ScriptPath = Resolve-Path $ScriptPath

# Lire la liste : ignorer les lignes vides et les commentaires (#)
$servers = Get-Content $VMList |
    Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*#' } |
    ForEach-Object { $_.Trim() } |
    Select-Object -Unique

if ($servers.Count -eq 0) {
    Write-Error "Aucun serveur trouvé dans $VMList"
    exit 1
}

# Créer le dossier de sortie local
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$OutputPath = Resolve-Path $OutputPath

#endregion

#region ── Affichage du résumé de lancement ────────────────────────────────────

Write-Host ""
Write-Host "══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  DEX Fleet - Génération sur $($servers.Count) serveur(s)"  -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Fichier VM     : $VMList"
Write-Host "  Script source  : $ScriptPath"
Write-Host "  Sortie locale  : $OutputPath"
Write-Host "  Parallélisme   : $ThrottleLimit sessions simultanées"
Write-Host "  Rédacteur      : $Redactor ($RedactorTitle)"
Write-Host "  Analyse reboots: $DaysToAnalyzeReboots jours"
Write-Host "  Transport      : $(if ($UseSSL) { 'HTTPS (5986)' } else { 'HTTP (5985)' })"
Write-Host ""
$servers | ForEach-Object { Write-Host "  [VM] $_" -ForegroundColor DarkCyan }
Write-Host ""

#endregion

#region ── Contenu du script à envoyer sur chaque VM ──────────────────────────

# On lit le contenu brut de Generate-DEX.ps1 pour le transmettre via ScriptBlock
$generateDexContent = Get-Content -Path $ScriptPath -Raw

#endregion

#region ── Rapport de synthèse ────────────────────────────────────────────────

$results = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()

#endregion

#region ── Exécution distante ─────────────────────────────────────────────────

# Paramètres communs de session PSRemoting
$sessionOpts = @{}
if ($UseSSL)      { $sessionOpts['UseSSL']      = $true }
if ($SkipCACheck) {
    $soOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $sessionOpts['SessionOption'] = $soOption
}
if ($Credential) { $sessionOpts['Credential'] = $Credential }

Write-Host "[*] Début de la génération..." -ForegroundColor Yellow
Write-Host ""

$startAll = Get-Date

# On traite en parallèle via des jobs (compatibles PS 5.1)
$jobs = @{}

foreach ($server in $servers) {
    Write-Host "  --> Lancement sur $server" -ForegroundColor DarkYellow

    $job = Start-Job -Name "DEX-$server" -ScriptBlock {
        param($server, $dexContent, $redactor, $redactorTitle, $days,
              $useSSL, $skipCA, $credential, $localOutput)

        $result = [PSCustomObject]@{
            Server    = $server
            Status    = "INCONNU"
            File      = ""
            Duration  = 0
            Error     = ""
        }
        $t0 = Get-Date

        try {
            # Construire les options de session
            $sopts = @{ ComputerName = $server; ErrorAction = 'Stop' }
            if ($credential)  { $sopts['Credential']    = $credential }
            if ($useSSL)      { $sopts['UseSSL']        = $true }
            if ($skipCA) {
                $sopts['SessionOption'] = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
            }

            $session = New-PSSession @sopts

            # Exécuter le script sur la VM distante (contenu injecté directement)
            $remoteFile = Invoke-Command -Session $session -ScriptBlock {
                param($content, $red, $redTitle, $days)

                # Écrire le script temporairement sur la VM distante
                $tmpScript = "$env:TEMP\Generate-DEX-tmp.ps1"
                $content | Out-File -FilePath $tmpScript -Encoding UTF8 -Force

                # Exécuter
                & $tmpScript `
                    -OutputPath   "C:\Temp\DEX" `
                    -Redactor     $red `
                    -RedactorTitle $redTitle `
                    -DaysToAnalyzeReboots $days

                # Retourner le chemin du fichier HTML le plus récent
                $latest = Get-ChildItem "C:\Temp\DEX" -Filter "DEX-Report-*.html" |
                          Sort-Object LastWriteTime -Descending |
                          Select-Object -First 1
                return $latest.FullName

            } -ArgumentList $content, $red, $redTitle, $days

            # Rapatrier le fichier HTML en local
            if ($remoteFile) {
                $localFile = Join-Path $localOutput ("$server-" + (Split-Path $remoteFile -Leaf))
                Copy-Item -FromSession $session -Path $remoteFile -Destination $localFile -Force
                $result.File   = $localFile
                $result.Status = "OK"
            } else {
                $result.Status = "WARN - Fichier HTML non trouvé sur la VM"
            }

            Remove-PSSession $session -ErrorAction SilentlyContinue

        } catch {
            $result.Status = "ERREUR"
            $result.Error  = $_.Exception.Message
        }

        $result.Duration = [math]::Round(((Get-Date) - $t0).TotalSeconds, 1)
        return $result

    } -ArgumentList $server, $generateDexContent, $Redactor, $RedactorTitle,
                    $DaysToAnalyzeReboots, $UseSSL.IsPresent, $SkipCACheck.IsPresent,
                    $Credential, $OutputPath

    $jobs[$server] = $job

    # Respecter le ThrottleLimit : attendre qu'une slot se libère
    while (($jobs.Values | Where-Object { $_.State -eq 'Running' }).Count -ge $ThrottleLimit) {
        Start-Sleep -Milliseconds 500
    }
}

# Attendre la fin de tous les jobs
Write-Host ""
Write-Host "[*] Attente de la fin de tous les jobs..." -ForegroundColor Yellow
$jobs.Values | Wait-Job | Out-Null

foreach ($srv in $jobs.Keys) {
    $jobResult = Receive-Job -Job $jobs[$srv]
    if ($jobResult) { $results.Add($jobResult) }
    Remove-Job -Job $jobs[$srv] -Force
}

#endregion

#region ── Rapport final ───────────────────────────────────────────────────────

$totalDuration = [math]::Round(((Get-Date) - $startAll).TotalSeconds, 1)
$ok    = @($results | Where-Object { $_.Status -eq "OK" })
$warn  = @($results | Where-Object { $_.Status -like "WARN*" })
$error_ = @($results | Where-Object { $_.Status -eq "ERREUR" })

Write-Host ""
Write-Host "══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  RAPPORT DE SYNTHÈSE DEX FLEET" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ("  Total     : {0} serveur(s)" -f $servers.Count)
Write-Host ("  Succès    : {0}" -f $ok.Count)    -ForegroundColor Green
Write-Host ("  Avertiss. : {0}" -f $warn.Count)  -ForegroundColor Yellow
Write-Host ("  Erreurs   : {0}" -f $error_.Count) -ForegroundColor Red
Write-Host ("  Durée     : {0}s" -f $totalDuration)
Write-Host ""

# Tableau détaillé
$colW = @(30, 10, 8, 50)
$header = "{0,-$($colW[0])} {1,-$($colW[1])} {2,-$($colW[2])} {3}" -f "Serveur","Statut","Durée(s)","Fichier / Erreur"
Write-Host $header -ForegroundColor White
Write-Host ("-" * 110)

foreach ($r in $results | Sort-Object Server) {
    $info  = if ($r.Status -eq "OK") { $r.File } else { $r.Error }
    $color = switch ($r.Status) {
        "OK"    { "Green"  }
        "ERREUR"{ "Red"    }
        default { "Yellow" }
    }
    $line = "{0,-$($colW[0])} {1,-$($colW[1])} {2,-$($colW[2])} {3}" -f $r.Server, $r.Status, $r.Duration, $info
    Write-Host $line -ForegroundColor $color
}

Write-Host ""
Write-Host "  Fichiers HTML disponibles dans : $OutputPath" -ForegroundColor Cyan
Write-Host ""

# Export CSV du rapport
$csvPath = Join-Path $OutputPath "DEX-Fleet-Summary-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$results | Select-Object Server, Status, Duration, File, Error |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Write-Host "  Rapport CSV : $csvPath" -ForegroundColor DarkCyan
Write-Host ""

#endregion

#region ── Envoi email via Power Automate (optionnel, contournement DLP) ───────

if ($SendEmail) {
    if ([string]::IsNullOrWhiteSpace($FlowWebhookUrl)) {
        Write-Warning "[-SendEmail] spécifié mais -FlowWebhookUrl est vide. Envoi ignoré."
    } elseif ($EmailTo.Count -eq 0) {
        Write-Warning "[-SendEmail] spécifié mais -EmailTo est vide. Envoi ignoré."
    } else {
        # Dot-sourcer Send-DEXReport.ps1 pour obtenir Send-DEXViaFlow
        $sendScript = Join-Path (Split-Path $ScriptPath -Parent) "Send-DEXReport.ps1"
        if (-not (Test-Path $sendScript)) {
            Write-Warning "Send-DEXReport.ps1 introuvable dans : $(Split-Path $ScriptPath -Parent)"
            Write-Warning "Placez Send-DEXReport.ps1 dans le même dossier que Generate-DEX.ps1."
        } else {
            . $sendScript

            Write-Host ""
            Write-Host "══════════════════════════════════════════════════" -ForegroundColor Cyan
            Write-Host "  ENVOI DES RAPPORTS PAR EMAIL (Power Automate)"   -ForegroundColor Cyan
            Write-Host "══════════════════════════════════════════════════" -ForegroundColor Cyan
            Write-Host "  Destinataires : $($EmailTo -join ' ; ')"
            if ($EmailCc) { Write-Host "  CC            : $EmailCc" }
            Write-Host ""

            $mailSent = 0; $mailErrors = 0

            foreach ($r in $ok) {
                if (-not $r.File -or -not (Test-Path $r.File)) {
                    Write-Warning "  [$($r.Server)] Fichier HTML introuvable, email ignoré."
                    $mailErrors++
                    continue
                }

                Write-Host "  --> Envoi pour $($r.Server)..." -NoNewline

                $res = Send-DEXViaFlow `
                    -WebhookUrl     $FlowWebhookUrl `
                    -To             $EmailTo `
                    -Cc             $EmailCc `
                    -AttachmentPath $r.File `
                    -ServerName     $r.Server

                if ($res.Success) {
                    Write-Host " OK" -ForegroundColor Green
                    $mailSent++
                } else {
                    Write-Host " ERREUR : $($res.Error)" -ForegroundColor Red
                    $mailErrors++
                }
            }

            Write-Host ""
            Write-Host ("  Emails envoyés : {0} / Erreurs : {1}" -f $mailSent, $mailErrors) `
                -ForegroundColor $(if ($mailErrors -eq 0) { "Green" } else { "Yellow" })
            Write-Host ""
        }
    }
}

#endregion
