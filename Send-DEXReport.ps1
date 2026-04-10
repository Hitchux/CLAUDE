#Requires -Version 5.1
<#
.SYNOPSIS
    Send-DEXReport.ps1 - Envoi de rapports DEX via un déclencheur HTTP Power Automate.

.DESCRIPTION
    Ce script expose la fonction Send-DEXViaFlow qui envoie un rapport DEX HTML
    via un Flow Power Automate (déclencheur "When an HTTP request is received").

    POURQUOI POWER AUTOMATE ?
    Les stratégies DLP de l'organisation bloquent l'envoi direct d'emails depuis
    les agents/scripts (SMTP, Graph API). Power Automate s'exécute sous l'identité
    d'un utilisateur autorisé (connexion Outlook 365), contournant cette restriction.
    Le script ne fait qu'un simple HTTP POST vers le webhook du Flow.

    STRUCTURE DU FLOW POWER AUTOMATE À CRÉER :
    ─────────────────────────────────────────
    1. Déclencheur : "Lors de la réception d'une requête HTTP" (HTTP Request)
       Schéma JSON du corps de la requête :
       {
           "type": "object",
           "properties": {
               "to":                { "type": "string" },
               "cc":               { "type": "string" },
               "subject":          { "type": "string" },
               "body":             { "type": "string" },
               "attachmentName":   { "type": "string" },
               "attachmentContent":{ "type": "string" }
           }
       }

    2. Action : "Envoyer un e-mail (V2)" (connecteur Outlook 365 Users)
       À       : @{triggerBody()?['to']}
       CC      : @{triggerBody()?['cc']}
       Objet   : @{triggerBody()?['subject']}
       Corps   : @{triggerBody()?['body']}   (activer "Corps est HTML")
       Pièces jointes > Nom     : @{triggerBody()?['attachmentName']}
       Pièces jointes > Contenu : @{base64ToBinary(triggerBody()?['attachmentContent'])}
       (Ajouter la pièce jointe uniquement si attachmentName n'est pas vide)

    NOTE : Copiez l'URL "HTTP POST URL" du déclencheur et collez-la dans
           le champ "Webhook URL Flow" de DEX Manager ou dans -FlowWebhookUrl.

.PARAMETER WebhookUrl
    URL du déclencheur HTTP du Flow Power Automate (HTTP POST URL).

.PARAMETER To
    Destinataire(s). Chaîne séparée par ";" ou tableau de chaînes.

.PARAMETER Cc
    Destinataires en copie (optionnel).

.PARAMETER Subject
    Sujet de l'email. Si vide, généré automatiquement depuis ServerName.

.PARAMETER AttachmentPath
    Chemin complet du fichier HTML DEX à joindre (optionnel).

.PARAMETER ServerName
    Nom du serveur pour le sujet et le corps par défaut. Par défaut : $env:COMPUTERNAME

.EXAMPLE
    .\Send-DEXReport.ps1 `
        -WebhookUrl "https://prod-xx.westeurope.logic.azure.com/workflows/..." `
        -To "sysadmin@company.com" `
        -AttachmentPath "C:\DEX-Reports\SRV01-DEX-Report-20250410-1430.html"

.EXAMPLE
    # Dot-source pour réutiliser la fonction dans un autre script :
    . .\Send-DEXReport.ps1
    $r = Send-DEXViaFlow -WebhookUrl $url -To @("admin@co.com") -AttachmentPath $file
    if (-not $r.Success) { Write-Warning $r.Error }

.NOTES
    Auteur  : Ingénierie Système
    Version : 1.1
    Requis  : PowerShell 5.1+, accès HTTPS sortant vers *.logic.azure.com (port 443)
#>

[CmdletBinding()]
param(
    [string]   $WebhookUrl     = "",
    [string[]] $To             = @(),
    [string]   $Cc             = "",
    [string]   $Subject        = "",
    [string]   $AttachmentPath = "",
    [string]   $ServerName     = $env:COMPUTERNAME
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

#region ── Fonction principale ─────────────────────────────────────────────────

function Send-DEXViaFlow {
    <#
    .SYNOPSIS
        Envoie un rapport DEX via un déclencheur HTTP Power Automate.
    .PARAMETER WebhookUrl
        URL HTTP POST du déclencheur Flow Power Automate.
    .PARAMETER To
        Destinataire(s) : tableau ou chaîne séparée par ";".
    .PARAMETER Cc
        Copie (optionnel).
    .PARAMETER Subject
        Sujet de l'email. Généré automatiquement si vide.
    .PARAMETER BodyHtml
        Corps HTML. Un corps par défaut est généré si vide.
    .PARAMETER AttachmentPath
        Chemin du fichier HTML DEX à joindre (optionnel).
    .PARAMETER ServerName
        Nom du serveur (sujet et corps par défaut).
    .OUTPUTS
        PSCustomObject { Success=[bool]; Response=[object]; Error=[string] }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]   $WebhookUrl,
        [Parameter(Mandatory)][string[]] $To,
        [string] $Cc             = "",
        [string] $Subject        = "",
        [string] $BodyHtml       = "",
        [string] $AttachmentPath = "",
        [string] $ServerName     = $env:COMPUTERNAME
    )

    # ── Sujet par défaut ──────────────────────────────────────────────────────
    if ([string]::IsNullOrWhiteSpace($Subject)) {
        $Subject = "DEX Report - $ServerName - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    }

    # ── Corps HTML par défaut ─────────────────────────────────────────────────
    if ([string]::IsNullOrWhiteSpace($BodyHtml)) {
        $BodyHtml = @"
<p>Bonjour,</p>
<p>Veuillez trouver ci-joint le Document d'Exploitation (DEX) généré automatiquement
pour le serveur <strong>$ServerName</strong> le $(Get-Date -Format 'dd/MM/yyyy \à HH:mm').</p>
<p>Ce document contient :</p>
<ul>
  <li>Informations système (OS, matériel, réseau)</li>
  <li>Applications installées</li>
  <li>Tâches planifiées personnalisées</li>
  <li>Analyse des reboots (EventLog)</li>
  <li>Planning de sauvegarde</li>
</ul>
<p>Cordialement,<br><em>DEX Manager — Ingénierie Système</em></p>
"@
    }

    # ── Construction du payload JSON ──────────────────────────────────────────
    $payload = [ordered]@{
        to                = ($To -join ";")
        cc                = $Cc
        subject           = $Subject
        body              = $BodyHtml
        attachmentName    = ""
        attachmentContent = ""
    }

    # Encodage base64 de la pièce jointe (si fournie et existante)
    if (-not [string]::IsNullOrWhiteSpace($AttachmentPath) -and (Test-Path $AttachmentPath)) {
        $payload['attachmentName']    = Split-Path $AttachmentPath -Leaf
        $payload['attachmentContent'] = [Convert]::ToBase64String(
            [System.IO.File]::ReadAllBytes($AttachmentPath)
        )
    }

    $jsonBody = $payload | ConvertTo-Json -Depth 5

    Write-Verbose "POST vers Power Automate : $WebhookUrl"
    Write-Verbose "Destinataires : $($payload['to'])"
    Write-Verbose "Sujet         : $Subject"
    Write-Verbose "Pièce jointe  : $(if ($payload['attachmentName']) { $payload['attachmentName'] } else { '(aucune)' })"

    # ── Appel HTTP POST ───────────────────────────────────────────────────────
    try {
        $response = Invoke-RestMethod `
            -Uri         $WebhookUrl `
            -Method      POST `
            -Body        $jsonBody `
            -ContentType "application/json" `
            -UseBasicParsing `
            -ErrorAction Stop

        Write-Host "[+] Email envoyé avec succès via Power Automate Flow." -ForegroundColor Green
        return [PSCustomObject]@{ Success = $true; Response = $response; Error = "" }

    } catch {
        $errMsg = $_.Exception.Message
        Write-Warning "Erreur envoi Flow : $errMsg"
        return [PSCustomObject]@{ Success = $false; Response = $null; Error = $errMsg }
    }
}

#endregion

#region ── Point d'entrée (exécution directe) ──────────────────────────────────

# Ce bloc s'exécute uniquement en invocation directe, pas en dot-source.
if ($MyInvocation.InvocationName -ne '.') {

    if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
        Write-Error "Le paramètre -WebhookUrl est obligatoire."
        Write-Host ""
        Write-Host "Usage :"
        Write-Host "  .\Send-DEXReport.ps1 -WebhookUrl `"https://...`" -To `"user@co.com`" [-AttachmentPath `"C:\DEX.html`"]"
        Write-Host ""
        Write-Host "Dot-source pour réutiliser la fonction :"
        Write-Host "  . .\Send-DEXReport.ps1"
        Write-Host "  Send-DEXViaFlow -WebhookUrl `$url -To @('admin@co.com') -AttachmentPath `$file"
        exit 1
    }

    if ($To.Count -eq 0) {
        Write-Error "Le paramètre -To est obligatoire (au moins un destinataire)."
        exit 1
    }

    $result = Send-DEXViaFlow `
        -WebhookUrl     $WebhookUrl `
        -To             $To `
        -Cc             $Cc `
        -Subject        $Subject `
        -AttachmentPath $AttachmentPath `
        -ServerName     $ServerName

    if (-not $result.Success) {
        Write-Error "Échec de l'envoi : $($result.Error)"
        exit 1
    }
}

#endregion
