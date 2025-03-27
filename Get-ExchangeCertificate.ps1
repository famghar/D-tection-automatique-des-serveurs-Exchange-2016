# Détection automatique des serveurs Exchange 2016
try {
    $servers = Get-ExchangeServer | Where-Object {$_.AdminDisplayVersion -like "Version 15.1*"} | Select-Object -ExpandProperty Name
    if ($servers -eq $null -or $servers.Count -eq 0) {
        Write-Warning "Aucun serveur Exchange 2016 n'a été trouvé dans l'organisation."
        exit 0
    }
} catch {
    Write-Error "Erreur lors de la récupération des serveurs Exchange : $($_.Exception.Message)"
    exit 1
}

# Collecte des informations sur les certificats et ajout d'une propriété pour la mise en évidence
$allCertificates = Invoke-Command -ComputerName $servers -ScriptBlock {
    Get-ExchangeCertificate | Select-Object FriendlyName, Subject, Issuer, NotBefore, NotAfter, Status, Thumbprint, Services,
        @{Name='Server';Expression={$env:COMPUTERNAME}},
        @{Name='Highlight';Expression={if ($_.NotAfter -lt (Get-Date).AddMonths(3)) {'Red'} else {'Black'}}}
} -ErrorAction Stop

# Fonction pour créer le rapport HTML groupé par serveur
function Build-HtmlReport {
    param(
        [array]$Certificates
    )

    $html = "<html><head><style>"
    $html += "table { border-collapse: collapse; width: 100%; }"
    $html += "th, td { border: 1px solid black; padding: 8px; text-align: left; }"
    $html += "th { background-color: #f2f2f2; }"
    $html += "</style></head><body>"
    $html += "<h2>Rapport des certificats Exchange</h2>"

    # Grouper les certificats par serveur
    $groupedCerts = $Certificates | Group-Object Server

    foreach ($group in $groupedCerts) {
        $html += "<h3>Serveur: $($group.Name)</h3>"
        $html += "<table border='1'>"
        $html += "<tr><th>Friendly Name</th><th>Subject</th><th>Issuer</th><th>Not Before</th><th>Not After</th><th>Status</th><th>Services</th></tr>"

        foreach ($cert in $group.Group) {
             $html += "<tr style='color:$($cert.Highlight)'>"
            $html += "<td>$($cert.FriendlyName)</td>"
            $html += "<td>$($cert.Subject)</td>"
            $html += "<td>$($cert.Issuer)</td>"
            $html += "<td>$($cert.NotBefore)</td>"
            $html += "<td>$($cert.NotAfter)</td>"
            $html += "<td>$($cert.Status)</td>"
            $html += "<td>$($cert.Services)</td>"
            $html += "</tr>"
        }
        $html += "</table><br/>"
    }

    $html += "</body></html>"
    return $html
}

# Création du rapport HTML
$htmlReport = Build-HtmlReport -Certificates $allCertificates


# Paramètres d'envoi d'e-mail
$smtpServer = "smtp.yourdomain.com"
$from = exchange-reports@yourdomain.com
$to = admin@yourdomain.com
$subject = "Rapport des certificats Exchange"

# Gestion sécurisée des informations d'identification (si nécessaire)
$credentialPath = "C:\path\to\credentials.xml"
if (Test-Path $credentialPath) {
    $credential = Import-CliXml -Path $credentialPath
} else {
    Write-Warning "Le fichier d'informations d'identification '$credentialPath' n'existe pas. L'authentification ne sera pas utilisée."
    $credential = $null
}

# Envoi de l'e-mail
try {
    if ($credential) {
        Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -Subject $subject -Body $htmlReport -BodyAsHtml -Credential $credential
    } else {
        Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -Subject $subject -Body $htmlReport -BodyAsHtml
    }
    Write-Host "Rapport envoyé avec succès."
} catch {
    Write-Error "Erreur lors de l'envoi de l'e-mail : $($_.Exception.Message)"
}

