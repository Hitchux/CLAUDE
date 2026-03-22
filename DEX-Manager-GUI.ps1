#Requires -Version 5.1
<#
.SYNOPSIS
    DEX-Manager-GUI.ps1 - Interface graphique de gestion DEX pour parc de serveurs Windows.

.DESCRIPTION
    GUI WinForms permettant de :
      - Charger une liste de VMs depuis vm.txt
      - Déployer et exécuter Generate-DEX.ps1 sur chaque machine
      - Collecter les rapports HTML en local via le bouton "Collecter"
      - Visualiser le statut en temps réel par VM
      - Ouvrir les rapports HTML directement depuis l'interface

.NOTES
    Requis : PowerShell 5.1+, .NET Framework 4.5+, WinRM actif sur les cibles
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "SilentlyContinue"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

#region ── Constantes & chemins ───────────────────────────────────────────────

$SCRIPT_DIR    = Split-Path -Parent $MyInvocation.MyCommand.Path
$DEX_SCRIPT    = Join-Path $SCRIPT_DIR "Generate-DEX.ps1"
$VM_LIST_FILE  = Join-Path $SCRIPT_DIR "vm.txt"
$LOCAL_OUTPUT  = Join-Path $SCRIPT_DIR "DEX-Reports"
$REMOTE_PATH   = "C:\Temp\DEX"

# Couleurs thème
$C_BG          = [System.Drawing.Color]::FromArgb(245, 247, 250)
$C_PANEL       = [System.Drawing.Color]::FromArgb(255, 255, 255)
$C_HEADER      = [System.Drawing.Color]::FromArgb(0,  51, 102)
$C_ACCENT      = [System.Drawing.Color]::FromArgb(72, 156, 223)
$C_BTN_DEPLOY  = [System.Drawing.Color]::FromArgb(0,  122, 162)
$C_BTN_COLLECT = [System.Drawing.Color]::FromArgb(40, 167,  69)
$C_BTN_OPEN    = [System.Drawing.Color]::FromArgb(255, 193,   7)
$C_BTN_CLEAR   = [System.Drawing.Color]::FromArgb(220,  53,  69)
$C_BTN_RELOAD  = [System.Drawing.Color]::FromArgb(108, 117, 125)

# Statuts VM
$STATUS = @{
    PENDING  = @{ Text="En attente";  Color=[System.Drawing.Color]::FromArgb(200,200,200) }
    RUNNING  = @{ Text="En cours..."; Color=[System.Drawing.Color]::FromArgb(255,200,  0) }
    OK       = @{ Text="OK";          Color=[System.Drawing.Color]::FromArgb( 40,167, 69) }
    WARN     = @{ Text="Avertiss.";   Color=[System.Drawing.Color]::FromArgb(255,140,  0) }
    ERROR    = @{ Text="Erreur";      Color=[System.Drawing.Color]::FromArgb(220, 53, 69) }
    COLLECTED= @{ Text="Collecté";    Color=[System.Drawing.Color]::FromArgb(  0,123,255) }
}

# Données par VM : [status, remoteFile, localFile, log]
$vmData = @{}

#endregion

#region ── Helpers UI ─────────────────────────────────────────────────────────

function New-Button {
    param([string]$Text, [int]$X, [int]$Y, [int]$W=140, [int]$H=36,
          [System.Drawing.Color]$BgColor=$C_BTN_DEPLOY)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text      = $Text
    $btn.Location  = New-Object System.Drawing.Point($X, $Y)
    $btn.Size      = New-Object System.Drawing.Size($W, $H)
    $btn.BackColor = $BgColor
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize = 0
    $btn.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$W=200, [int]$H=22,
          [System.Drawing.Color]$FgColor=[System.Drawing.Color]::Black,
          [float]$FontSize=9, [System.Drawing.FontStyle]$Style=[System.Drawing.FontStyle]::Regular)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Text
    $lbl.Location  = New-Object System.Drawing.Point($X, $Y)
    $lbl.Size      = New-Object System.Drawing.Size($W, $H)
    $lbl.ForeColor = $FgColor
    $lbl.Font      = New-Object System.Drawing.Font("Segoe UI", $FontSize, $Style)
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    return $lbl
}

function Log-Msg {
    param([string]$Msg, [string]$Level="INFO")
    $ts    = Get-Date -Format "HH:mm:ss"
    $color = switch ($Level) {
        "OK"    { ">>" } "ERR"  { "!!" } "WARN" { "!>" } default { "--" }
    }
    $line = "[$ts] $color $Msg"
    $script:txtLog.AppendText("$line`r`n")
    $script:txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-VMRow {
    param([string]$Server, [string]$StatusKey)
    $row = $script:lvVMs.Items | Where-Object { $_.Text -eq $Server }
    if (-not $row) { return }
    $st = $STATUS[$StatusKey]
    $row.SubItems[1].Text = $st.Text
    $row.UseItemStyleForSubItems = $false
    $row.SubItems[1].BackColor   = $st.Color
    $row.SubItems[1].ForeColor   = if ($StatusKey -in @("RUNNING","PENDING")) {
        [System.Drawing.Color]::Black } else { [System.Drawing.Color]::White }
    [System.Windows.Forms.Application]::DoEvents()
}

function Refresh-VMList {
    $script:lvVMs.Items.Clear()
    $script:vmData.Clear()

    if (-not (Test-Path $script:VM_LIST_FILE)) {
        Log-Msg "Fichier vm.txt introuvable : $($script:VM_LIST_FILE)" "ERR"
        return
    }
    $servers = Get-Content $script:VM_LIST_FILE |
        Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*#' } |
        ForEach-Object { $_.Trim() } | Select-Object -Unique

    foreach ($srv in $servers) {
        $item = New-Object System.Windows.Forms.ListViewItem($srv)
        $item.SubItems.Add("En attente") | Out-Null
        $item.SubItems.Add("")           | Out-Null   # Fichier distant
        $item.SubItems.Add("")           | Out-Null   # Fichier local
        $script:lvVMs.Items.Add($item) | Out-Null
        $script:vmData[$srv] = @{ Status="PENDING"; RemoteFile=""; LocalFile=""; Log="" }
    }
    Log-Msg "$($servers.Count) serveur(s) chargé(s) depuis vm.txt" "OK"
}

#endregion

#region ── Actions métier ─────────────────────────────────────────────────────

function Get-SessionOpts {
    $opts = @{}
    if ($script:chkSSL.Checked) {
        $opts['UseSSL']        = $true
        $opts['SessionOption'] = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    }
    $user = $script:txtUser.Text.Trim()
    $pass = $script:txtPass.Text
    if ($user -ne "") {
        $secPass = ConvertTo-SecureString $pass -AsPlainText -Force
        $opts['Credential'] = New-Object System.Management.Automation.PSCredential($user, $secPass)
    }
    return $opts
}

function Deploy-And-Execute {
    param([string[]]$Servers)

    if (-not (Test-Path $script:DEX_SCRIPT)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Generate-DEX.ps1 introuvable :`n$($script:DEX_SCRIPT)",
            "Erreur", "OK", "Error") | Out-Null
        return
    }

    $dexContent  = Get-Content $script:DEX_SCRIPT -Raw
    $redactor    = $script:txtRedactor.Text.Trim()
    $redTitle    = $script:txtRedTitle.Text.Trim()
    $days        = [int]$script:numDays.Value
    $sopts       = Get-SessionOpts
    $throttle    = [int]$script:numThrottle.Value

    $script:btnDeploy.Enabled  = $false
    $script:btnCollect.Enabled = $false

    Log-Msg "Déploiement sur $($Servers.Count) serveur(s) [parallélisme=$throttle]..." "INFO"

    $jobs    = @{}
    $pending = [System.Collections.Queue]::new()
    $Servers | ForEach-Object { $pending.Enqueue($_) }

    while ($pending.Count -gt 0 -or $jobs.Count -gt 0) {

        # Lancer de nouveaux jobs si des slots sont disponibles
        while ($pending.Count -gt 0 -and $jobs.Count -lt $throttle) {
            $srv = $pending.Dequeue()
            Update-VMRow $srv "RUNNING"
            Log-Msg "[$srv] Connexion WinRM en cours..."

            $job = Start-Job -ScriptBlock {
                param($srv, $content, $red, $redTitle, $days, $sopts, $remotePath)
                $r = [PSCustomObject]@{
                    Server="$srv"; Status="ERROR"; RemoteFile=""; Error="" }
                try {
                    $s = New-PSSession -ComputerName $srv @sopts -ErrorAction Stop
                    $rf = Invoke-Command -Session $s -ScriptBlock {
                        param($c, $r, $rt, $d, $rp)
                        $tmp = "$env:TEMP\Generate-DEX-tmp.ps1"
                        $c | Out-File $tmp -Encoding UTF8 -Force
                        & $tmp -OutputPath $rp -Redactor $r -RedactorTitle $rt `
                               -DaysToAnalyzeReboots $d
                        $f = Get-ChildItem $rp -Filter "DEX-Report-*.html" |
                             Sort-Object LastWriteTime -Descending |
                             Select-Object -First 1
                        return $f.FullName
                    } -ArgumentList $content,$red,$redTitle,$days,$remotePath
                    if ($rf) { $r.Status="OK"; $r.RemoteFile=$rf }
                    else     { $r.Status="WARN"; $r.Error="Fichier HTML non trouvé" }
                    Remove-PSSession $s -EA SilentlyContinue
                } catch { $r.Error=$_.Exception.Message }
                return $r
            } -ArgumentList $srv, $dexContent, $redactor, $redTitle, $days, $sopts, $REMOTE_PATH

            $jobs[$srv] = $job
        }

        # Vérifier les jobs terminés
        $done = $jobs.GetEnumerator() | Where-Object { $_.Value.State -ne 'Running' }
        foreach ($entry in $done) {
            $srv = $entry.Key
            $res = Receive-Job -Job $entry.Value
            Remove-Job -Job $entry.Value -Force
            $jobs.Remove($srv)

            if ($res) {
                if ($res.Status -eq "OK") {
                    $script:vmData[$srv]['Status']     = "OK"
                    $script:vmData[$srv]['RemoteFile'] = $res.RemoteFile
                    Update-VMRow $srv "OK"
                    $row = $script:lvVMs.Items | Where-Object { $_.Text -eq $srv }
                    if ($row) { $row.SubItems[2].Text = $res.RemoteFile }
                    Log-Msg "[$srv] DEX genere : $($res.RemoteFile)" "OK"
                } elseif ($res.Status -eq "WARN") {
                    $script:vmData[$srv]['Status'] = "WARN"
                    Update-VMRow $srv "WARN"
                    Log-Msg "[$srv] AVERT. $($res.Error)" "WARN"
                } else {
                    $script:vmData[$srv]['Status'] = "ERROR"
                    Update-VMRow $srv "ERROR"
                    Log-Msg "[$srv] ERREUR $($res.Error)" "ERR"
                }
            }
        }

        if ($jobs.Count -gt 0 -or $pending.Count -gt 0) {
            Start-Sleep -Milliseconds 600
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    Log-Msg "Déploiement terminé." "OK"
    $script:btnDeploy.Enabled  = $true
    $script:btnCollect.Enabled = $true
}

function Collect-Reports {
    param([string[]]$Servers)

    if (-not (Test-Path $LOCAL_OUTPUT)) {
        New-Item -ItemType Directory -Path $LOCAL_OUTPUT -Force | Out-Null
    }

    $sopts    = Get-SessionOpts
    $throttle = [int]$script:numThrottle.Value

    $script:btnDeploy.Enabled  = $false
    $script:btnCollect.Enabled = $false

    Log-Msg "Collecte des rapports HTML sur $($Servers.Count) serveur(s)..." "INFO"

    $jobs    = @{}
    $pending = [System.Collections.Queue]::new()
    $Servers | ForEach-Object {
        if ($script:vmData[$_]['Status'] -in @("OK","WARN","COLLECTED")) {
            $pending.Enqueue($_)
        } else {
            Log-Msg "[$_] Ignoré (statut : $($script:vmData[$_]['Status']))" "WARN"
        }
    }

    if ($pending.Count -eq 0) {
        Log-Msg "Aucun serveur prêt à être collecté. Lancez d'abord le déploiement." "WARN"
        $script:btnDeploy.Enabled  = $true
        $script:btnCollect.Enabled = $true
        return
    }

    $localOut = $LOCAL_OUTPUT

    while ($pending.Count -gt 0 -or $jobs.Count -gt 0) {

        while ($pending.Count -gt 0 -and $jobs.Count -lt $throttle) {
            $srv        = $pending.Dequeue()
            $remotefile = $script:vmData[$srv]['RemoteFile']
            Log-Msg "[$srv] Collecte de $remotefile..."

            $job = Start-Job -ScriptBlock {
                param($srv, $remoteFile, $localOut, $sopts)
                $r = [PSCustomObject]@{ Server=$srv; Status="ERROR"; LocalFile=""; Error="" }
                try {
                    $s = New-PSSession -ComputerName $srv @sopts -ErrorAction Stop
                    # Si on n'a pas le chemin exact, chercher le plus récent
                    if (-not $remoteFile) {
                        $remoteFile = Invoke-Command -Session $s -ScriptBlock {
                            $f = Get-ChildItem "C:\Temp\DEX" -Filter "DEX-Report-*.html" -EA SilentlyContinue |
                                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
                            return $f.FullName
                        }
                    }
                    if ($remoteFile) {
                        $localFile = Join-Path $localOut ("$srv-" + (Split-Path $remoteFile -Leaf))
                        Copy-Item -FromSession $s -Path $remoteFile -Destination $localFile -Force
                        $r.Status    = "OK"
                        $r.LocalFile = $localFile
                    } else {
                        $r.Error = "Aucun rapport HTML trouvé sur la VM"
                    }
                    Remove-PSSession $s -EA SilentlyContinue
                } catch { $r.Error = $_.Exception.Message }
                return $r
            } -ArgumentList $srv, $remotefile, $localOut, $sopts

            $jobs[$srv] = $job
        }

        $done = $jobs.GetEnumerator() | Where-Object { $_.Value.State -ne 'Running' }
        foreach ($entry in $done) {
            $srv = $entry.Key
            $res = Receive-Job -Job $entry.Value
            Remove-Job -Job $entry.Value -Force
            $jobs.Remove($srv)

            if ($res -and $res.Status -eq "OK") {
                $script:vmData[$srv]['Status']    = "COLLECTED"
                $script:vmData[$srv]['LocalFile'] = $res.LocalFile
                Update-VMRow $srv "COLLECTED"
                $row = $script:lvVMs.Items | Where-Object { $_.Text -eq $srv }
                if ($row) { $row.SubItems[3].Text = $res.LocalFile }
                Log-Msg "[$srv] Collecte OK --> $($res.LocalFile)" "OK"
            } else {
                $err = if ($res) { $res.Error } else { "Résultat vide" }
                $script:vmData[$srv]['Status'] = "ERROR"
                Update-VMRow $srv "ERROR"
                Log-Msg "[$srv] ERREUR $err" "ERR"
            }
        }

        if ($jobs.Count -gt 0 -or $pending.Count -gt 0) {
            Start-Sleep -Milliseconds 600
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    Log-Msg "Collecte terminée. Rapports dans : $LOCAL_OUTPUT" "OK"
    $script:btnDeploy.Enabled  = $true
    $script:btnCollect.Enabled = $true
}

#endregion

#region ── Construction du formulaire principal ───────────────────────────────

$form                  = New-Object System.Windows.Forms.Form
$form.Text             = "DEX Manager - Gestion du parc serveurs Windows"
$form.Size             = New-Object System.Drawing.Size(1100, 780)
$form.StartPosition    = "CenterScreen"
$form.BackColor        = $C_BG
$form.MinimumSize      = New-Object System.Drawing.Size(900, 650)
$form.Font             = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Icon             = [System.Drawing.SystemIcons]::Shield

# ── Bandeau titre ────────────────────────────────────────────────────────────
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock      = "Top"
$pnlHeader.Height    = 60
$pnlHeader.BackColor = $C_HEADER
$form.Controls.Add($pnlHeader)

$lblTitle = New-Label "[*] DEX Manager" 15 8 500 26 ([System.Drawing.Color]::White) 14 Bold
$pnlHeader.Controls.Add($lblTitle)
$lblSub = New-Label "Generation automatique de Documents d'Exploitation - Windows Server 2016+" 15 36 700 18 ([System.Drawing.Color]::FromArgb(180,210,240)) 9
$pnlHeader.Controls.Add($lblSub)

# ── Panneau gauche : configuration ──────────────────────────────────────────
$pnlLeft = New-Object System.Windows.Forms.Panel
$pnlLeft.Location  = New-Object System.Drawing.Point(0, 60)
$pnlLeft.Size      = New-Object System.Drawing.Size(260, 680)
$pnlLeft.BackColor = $C_PANEL
$pnlLeft.BorderStyle = "FixedSingle"
$form.Controls.Add($pnlLeft)

$y = 10
$pnlLeft.Controls.Add((New-Label "CONFIGURATION" 10 $y 230 20 $C_ACCENT 8 Bold))
$y += 28

# Fichier VM
$pnlLeft.Controls.Add((New-Label "Fichier VM list :" 10 $y 220 18 ([System.Drawing.Color]::FromArgb(80,80,80))))
$y += 20
$script:txtVMFile = New-Object System.Windows.Forms.TextBox
$script:txtVMFile.Location = New-Object System.Drawing.Point(10, $y)
$script:txtVMFile.Size     = New-Object System.Drawing.Size(190, 22)
$script:txtVMFile.Text     = $VM_LIST_FILE
$pnlLeft.Controls.Add($script:txtVMFile)
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location  = New-Object System.Drawing.Point(205, $y)
$btnBrowse.Size      = New-Object System.Drawing.Size(45, 22)
$btnBrowse.Text      = "..."
$btnBrowse.FlatStyle = "Flat"
$btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(220,220,220)
$pnlLeft.Controls.Add($btnBrowse)
$y += 30

# Credentials
$pnlLeft.Controls.Add((New-Label "Compte (domaine\user) :" 10 $y 220 18 ([System.Drawing.Color]::FromArgb(80,80,80))))
$y += 20
$script:txtUser = New-Object System.Windows.Forms.TextBox
$script:txtUser.Location    = New-Object System.Drawing.Point(10, $y)
$script:txtUser.Size        = New-Object System.Drawing.Size(240, 22)
$script:txtUser.PlaceholderText = "Laisser vide = credentials courants"
$pnlLeft.Controls.Add($script:txtUser)
$y += 28

$pnlLeft.Controls.Add((New-Label "Mot de passe :" 10 $y 220 18 ([System.Drawing.Color]::FromArgb(80,80,80))))
$y += 20
$script:txtPass = New-Object System.Windows.Forms.TextBox
$script:txtPass.Location     = New-Object System.Drawing.Point(10, $y)
$script:txtPass.Size         = New-Object System.Drawing.Size(240, 22)
$script:txtPass.PasswordChar = '*'
$pnlLeft.Controls.Add($script:txtPass)
$y += 30

# SSL
$script:chkSSL = New-Object System.Windows.Forms.CheckBox
$script:chkSSL.Text     = "WinRM over HTTPS (port 5986)"
$script:chkSSL.Location = New-Object System.Drawing.Point(10, $y)
$script:chkSSL.Size     = New-Object System.Drawing.Size(240, 22)
$pnlLeft.Controls.Add($script:chkSSL)
$y += 30

# Séparateur
$sep1 = New-Object System.Windows.Forms.Label
$sep1.BorderStyle = "Fixed3D"
$sep1.Location = New-Object System.Drawing.Point(10, $y)
$sep1.Size     = New-Object System.Drawing.Size(240, 2)
$pnlLeft.Controls.Add($sep1)
$y += 10

# Rédacteur
$pnlLeft.Controls.Add((New-Label "Rédacteur :" 10 $y 220 18 ([System.Drawing.Color]::FromArgb(80,80,80))))
$y += 20
$script:txtRedactor = New-Object System.Windows.Forms.TextBox
$script:txtRedactor.Location = New-Object System.Drawing.Point(10, $y)
$script:txtRedactor.Size     = New-Object System.Drawing.Size(240, 22)
$script:txtRedactor.Text     = $env:USERNAME
$pnlLeft.Controls.Add($script:txtRedactor)
$y += 28

$pnlLeft.Controls.Add((New-Label "Fonction :" 10 $y 220 18 ([System.Drawing.Color]::FromArgb(80,80,80))))
$y += 20
$script:txtRedTitle = New-Object System.Windows.Forms.TextBox
$script:txtRedTitle.Location = New-Object System.Drawing.Point(10, $y)
$script:txtRedTitle.Size     = New-Object System.Drawing.Size(240, 22)
$script:txtRedTitle.Text     = "System Engineer"
$pnlLeft.Controls.Add($script:txtRedTitle)
$y += 30

# Jours analyse reboots
$pnlLeft.Controls.Add((New-Label "Jours analyse reboots :" 10 $y 220 18 ([System.Drawing.Color]::FromArgb(80,80,80))))
$y += 20
$script:numDays = New-Object System.Windows.Forms.NumericUpDown
$script:numDays.Location = New-Object System.Drawing.Point(10, $y)
$script:numDays.Size     = New-Object System.Drawing.Size(80, 22)
$script:numDays.Minimum  = 7
$script:numDays.Maximum  = 365
$script:numDays.Value    = 90
$pnlLeft.Controls.Add($script:numDays)
$y += 30

# Parallélisme
$pnlLeft.Controls.Add((New-Label "Sessions parallèles :" 10 $y 220 18 ([System.Drawing.Color]::FromArgb(80,80,80))))
$y += 20
$script:numThrottle = New-Object System.Windows.Forms.NumericUpDown
$script:numThrottle.Location = New-Object System.Drawing.Point(10, $y)
$script:numThrottle.Size     = New-Object System.Drawing.Size(80, 22)
$script:numThrottle.Minimum  = 1
$script:numThrottle.Maximum  = 20
$script:numThrottle.Value    = 5
$pnlLeft.Controls.Add($script:numThrottle)
$y += 40

# Séparateur
$sep2 = New-Object System.Windows.Forms.Label
$sep2.BorderStyle = "Fixed3D"
$sep2.Location = New-Object System.Drawing.Point(10, $y)
$sep2.Size     = New-Object System.Drawing.Size(240, 2)
$pnlLeft.Controls.Add($sep2)
$y += 10

# Dossier de sortie local
$pnlLeft.Controls.Add((New-Label "Dossier sortie local :" 10 $y 220 18 ([System.Drawing.Color]::FromArgb(80,80,80))))
$y += 20
$script:txtOutput = New-Object System.Windows.Forms.TextBox
$script:txtOutput.Location = New-Object System.Drawing.Point(10, $y)
$script:txtOutput.Size     = New-Object System.Drawing.Size(190, 22)
$script:txtOutput.Text     = $LOCAL_OUTPUT
$pnlLeft.Controls.Add($script:txtOutput)
$btnBrowseOut = New-Object System.Windows.Forms.Button
$btnBrowseOut.Location  = New-Object System.Drawing.Point(205, $y)
$btnBrowseOut.Size      = New-Object System.Drawing.Size(45, 22)
$btnBrowseOut.Text      = "..."
$btnBrowseOut.FlatStyle = "Flat"
$btnBrowseOut.BackColor = [System.Drawing.Color]::FromArgb(220,220,220)
$pnlLeft.Controls.Add($btnBrowseOut)

# ── Zone centrale : liste des VMs ────────────────────────────────────────────
$pnlCenter = New-Object System.Windows.Forms.Panel
$pnlCenter.Location  = New-Object System.Drawing.Point(260, 60)
$pnlCenter.Size      = New-Object System.Drawing.Size(840, 430)
$pnlCenter.BackColor = $C_BG
$form.Controls.Add($pnlCenter)

# Toolbar boutons
$script:btnReload = New-Button "[R] Recharger"   10  8 130 34 $C_BTN_RELOAD
$script:btnDeploy = New-Button "[>] Deployer"    150  8 130 34 $C_BTN_DEPLOY
$script:btnCollect= New-Button "[v] Collecter"   290  8 130 34 $C_BTN_COLLECT
$script:btnOpen   = New-Button "[W] Ouvrir DEX"  430  8 130 34 $C_BTN_OPEN
$script:btnClear  = New-Button "[X] Effacer log" 660  8 130 34 $C_BTN_CLEAR
$pnlCenter.Controls.AddRange(@($script:btnReload, $script:btnDeploy, $script:btnCollect,
                                $script:btnOpen,  $script:btnClear))

# Légende statuts
$xLeg = 570; $yLeg = 12
foreach ($key in @("PENDING","RUNNING","OK","WARN","ERROR","COLLECTED")) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $STATUS[$key].Text
    $lbl.Location  = New-Object System.Drawing.Point($xLeg, $yLeg)
    $lbl.Size      = New-Object System.Drawing.Size(80, 16)
    $lbl.BackColor = $STATUS[$key].Color
    $lbl.ForeColor = if ($key -in @("PENDING","RUNNING")) { [System.Drawing.Color]::Black } else { [System.Drawing.Color]::White }
    $lbl.Font      = New-Object System.Drawing.Font("Segoe UI", 7)
    $lbl.TextAlign = "MiddleCenter"
    $pnlCenter.Controls.Add($lbl)
    $yLeg += 19
    if ($key -eq "ERROR") { $xLeg = 655; $yLeg = 12 }
}

# ListView VMs
$script:lvVMs = New-Object System.Windows.Forms.ListView
$script:lvVMs.Location         = New-Object System.Drawing.Point(10, 52)
$script:lvVMs.Size             = New-Object System.Drawing.Size(820, 365)
$script:lvVMs.View             = "Details"
$script:lvVMs.FullRowSelect    = $true
$script:lvVMs.GridLines        = $true
$script:lvVMs.MultiSelect      = $true
$script:lvVMs.BackColor        = $C_PANEL
$script:lvVMs.Font             = New-Object System.Drawing.Font("Consolas", 9)
$script:lvVMs.BorderStyle      = "FixedSingle"
$null = $script:lvVMs.Columns.Add("Serveur",            180)
$null = $script:lvVMs.Columns.Add("Statut",             100)
$null = $script:lvVMs.Columns.Add("Rapport (distant)",  250)
$null = $script:lvVMs.Columns.Add("Rapport (local)",    280)
$pnlCenter.Controls.Add($script:lvVMs)

# ── Zone basse : log ─────────────────────────────────────────────────────────
$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Location  = New-Object System.Drawing.Point(260, 490)
$pnlLog.Size      = New-Object System.Drawing.Size(840, 250)
$pnlLog.BackColor = $C_PANEL
$pnlLog.BorderStyle = "FixedSingle"
$form.Controls.Add($pnlLog)

$lblLog = New-Label "Journal d'activité" 8 4 200 18 $C_ACCENT 8 Bold
$pnlLog.Controls.Add($lblLog)

$script:txtLog = New-Object System.Windows.Forms.RichTextBox
$script:txtLog.Location    = New-Object System.Drawing.Point(8, 24)
$script:txtLog.Size        = New-Object System.Drawing.Size(824, 218)
$script:txtLog.BackColor   = [System.Drawing.Color]::FromArgb(20,20,30)
$script:txtLog.ForeColor   = [System.Drawing.Color]::FromArgb(180,220,180)
$script:txtLog.Font        = New-Object System.Drawing.Font("Consolas", 8.5)
$script:txtLog.ReadOnly    = $true
$script:txtLog.ScrollBars  = "Vertical"
$script:txtLog.BorderStyle = "None"
$pnlLog.Controls.Add($script:txtLog)

# ── Barre de statut ──────────────────────────────────────────────────────────
$statusBar = New-Object System.Windows.Forms.StatusStrip
$script:lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblStatus.Text  = "Prêt"
$script:pbStatus  = New-Object System.Windows.Forms.ToolStripProgressBar
$script:pbStatus.Width  = 150
$script:pbStatus.Style  = "Marquee"
$script:pbStatus.Visible= $false
$statusBar.Items.AddRange(@($script:lblStatus, $script:pbStatus))
$form.Controls.Add($statusBar)

#endregion

#region ── Événements ─────────────────────────────────────────────────────────

# Parcourir fichier VM list
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Fichiers texte (*.txt)|*.txt|Tous (*.*)|*.*"
    $dlg.InitialDirectory = $SCRIPT_DIR
    if ($dlg.ShowDialog() -eq "OK") {
        $script:txtVMFile.Text = $dlg.FileName
        $script:VM_LIST_FILE   = $dlg.FileName
        Refresh-VMList
    }
})

# Parcourir dossier sortie
$btnBrowseOut.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description  = "Choisir le dossier de sortie local pour les rapports HTML"
    $dlg.SelectedPath = $script:txtOutput.Text
    if ($dlg.ShowDialog() -eq "OK") {
        $script:txtOutput.Text = $dlg.SelectedPath
        $script:LOCAL_OUTPUT   = $dlg.SelectedPath
    }
})

# Recharger VM list
$script:btnReload.Add_Click({
    $script:VM_LIST_FILE = $script:txtVMFile.Text
    Refresh-VMList
})

# Déployer & Exécuter
$script:btnDeploy.Add_Click({
    $selected = $script:lvVMs.SelectedItems
    $targets  = if ($selected.Count -gt 0) {
        $selected | ForEach-Object { $_.Text }
    } else {
        $script:lvVMs.Items | ForEach-Object { $_.Text }
    }
    if ($targets.Count -eq 0) { Log-Msg "Aucun serveur dans la liste." "WARN"; return }

    $script:lblStatus.Text   = "Déploiement en cours..."
    $script:pbStatus.Visible = $true
    [System.Windows.Forms.Application]::DoEvents()

    Deploy-And-Execute -Servers $targets

    $script:lblStatus.Text   = "Déploiement terminé"
    $script:pbStatus.Visible = $false
})

# Collecter rapports
$script:btnCollect.Add_Click({
    $selected = $script:lvVMs.SelectedItems
    $targets  = if ($selected.Count -gt 0) {
        $selected | ForEach-Object { $_.Text }
    } else {
        $script:lvVMs.Items | ForEach-Object { $_.Text }
    }
    if ($targets.Count -eq 0) { Log-Msg "Aucun serveur dans la liste." "WARN"; return }

    $script:LOCAL_OUTPUT     = $script:txtOutput.Text
    $script:lblStatus.Text   = "Collecte en cours..."
    $script:pbStatus.Visible = $true
    [System.Windows.Forms.Application]::DoEvents()

    Collect-Reports -Servers $targets

    $script:lblStatus.Text   = "Collecte terminée"
    $script:pbStatus.Visible = $false
})

# Ouvrir le rapport sélectionné
$script:btnOpen.Add_Click({
    $sel = $script:lvVMs.SelectedItems | Select-Object -First 1
    if (-not $sel) { Log-Msg "Sélectionnez un serveur dans la liste." "WARN"; return }
    $srv       = $sel.Text
    $localFile = $script:vmData[$srv]['LocalFile']
    if ($localFile -and (Test-Path $localFile)) {
        Start-Process $localFile
        Log-Msg "Ouverture : $localFile" "OK"
    } else {
        # Chercher dans le dossier local par nom de serveur
        $found = Get-ChildItem $script:txtOutput.Text -Filter "$srv-DEX-Report-*.html" -EA SilentlyContinue |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($found) {
            Start-Process $found.FullName
            Log-Msg "Ouverture : $($found.FullName)" "OK"
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Rapport non trouvé pour $srv.`nLancez d'abord Déployer puis Collecter.",
                "Rapport introuvable", "OK", "Warning") | Out-Null
        }
    }
})

# Double-clic sur une VM : ouvre le rapport si disponible
$script:lvVMs.Add_DoubleClick({
    $script:btnOpen.PerformClick()
})

# Effacer le log
$script:btnClear.Add_Click({
    $script:txtLog.Clear()
})

# Resize : adapter les panneaux
$form.Add_Resize({
    $w = $form.ClientSize.Width
    $h = $form.ClientSize.Height
    $pnlCenter.Width  = $w - 260
    $pnlLog.Width     = $w - 260
    $pnlLog.Top       = $h - 260
    $pnlLeft.Height   = $h - 60
    $script:lvVMs.Size = New-Object System.Drawing.Size(($w - 280), 365)
    $script:txtLog.Size= New-Object System.Drawing.Size(($w - 276), 218)
})

# Fermeture propre
$form.Add_FormClosing({
    Get-Job | Remove-Job -Force -EA SilentlyContinue
})

#endregion

#region ── Initialisation ─────────────────────────────────────────────────────

Refresh-VMList
Log-Msg "DEX Manager démarré. Script source : $DEX_SCRIPT" "INFO"
Log-Msg "Raccourcis : Ctrl+clic ou Shift+clic pour sélection multiple dans la liste." "INFO"
Log-Msg "Double-clic sur une VM pour ouvrir son rapport." "INFO"

#endregion

# Lancement de la fenêtre
[System.Windows.Forms.Application]::Run($form)
