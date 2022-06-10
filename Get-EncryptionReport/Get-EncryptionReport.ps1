<#
.SYNOPSIS
  Name: Get-EncryptionReport.ps1
  Create report for in- and outbound E-Mails which were encrypted.
  Requires NSP version 13.2.20171.1151 or later.

.DESCRIPTION
 This script generates a report over the count of encrypted and non encrypted mails. It seperates results in both directions and timespan (last 30 days and 30 days before that timespan).

.PARAMETER smtphost
  Specifies the Hostname or IP of the mailserver the report will be sent to.
  E.g.:mailserver.example.com

.PARAMETER ReportSender
  Specifies the mailaddress used as sender for the report.
  E.g.:server@example.com

.PARAMETER ReportRecipient
  Specifies the recipient address for the report.
  E.g.: alice@example.com

.OUTPUTS
  Sends an email to the specified recipient with an html-report.

.NOTES
  Version:        1.1.0
  Author:         Finn Schulte
  Creation Date:  2022-06-07
  Purpose/Change: added v14 support
  
.LINK
  https://www.nospamproxy.de
  https://forum.nospamproxy.com
  https://github.com/noSpamProxy

.EXAMPLE
  .\Get-EncryptionReport.ps1 -smtphost "mailserver.company" -ReportSender "mailgateway@company" -ReportRecipient "admin@company"
#>

param (
  [Parameter(Mandatory=$true)][string] $smtphost,
  [Parameter(Mandatory=$true)][string] $ReportSender,
  [Parameter(Mandatory=$true)][string] $ReportRecipient,
  # only needed for v14 with enabled provider mode
	[Parameter(Mandatory=$false)][string] $TenantPrimaryDomain
)

$nspVersion = (Get-ItemProperty -Path HKLM:\SOFTWARE\NoSpamProxy\Components -ErrorAction SilentlyContinue).'Intranet Role'
if ($nspVersion -gt '14.0') {
	try {
		Connect-Nsp -IgnoreServerCertificateErrors -ErrorAction Stop
	} catch {
		$e = $_
		Write-Warning "Not possible to connect with the NoSpamProxy. Please check the error message below."
		$e |Format-List * -Force
		EXIT
	}
	if ($(Get-NspIsProviderModeEnabled) -eq $true) {
		if ($null -eq $TenantPrimaryDomain -OR $TenantPrimaryDomain -eq "") {
			Write-Host "Please provide a TenantPrimaryDomain to run this script with NoSpamProxy v14 in provider mode."
			EXIT
		} else {
			# NSP v14 has a new authentication mechanism, Connect-Nsp is required to authenticate properly
			# -IgnoreServerCertificateErrors allows the usage of self-signed certificates
			Connect-Nsp -IgnoreServerCertificateErrors -PrimaryDomain $TenantPrimaryDomain
		}
	}
}

$outboundencrypted = 0
$outboundnonencrypted = 0
$inboundencrypted = 0
$inboundnonencrypted = 0

for($x=0;$x -lt 30;$x++){

$MessageTracks = Get-NspMessageTrack -From (Get-Date).AddDays(-1-$x) -To (Get-Date).AddDays(0-$x) 
   
    Foreach ($MessageTrack in $MessageTracks)
    {
        if ($MessageTrack.WasReceivedFromRelayServer)
        {
            if ($MessageTrack.Encrypted -notlike 'None')
            {
                $outboundencrypted+=1
            }
            else
            {
                $outboundnonencrypted+=1
            }
        }
        else
        {
            if ($MessageTrack.Encrypted -notlike 'None')
            {
                $inboundencrypted+=1
            }
            else
            {
                $inboundnonencrypted+=1
            }
        }
    }
Write-Host ""
}
 $global:htmlout = "<html>
		<head>
			<title>Auswertung über verschlüsselte E-Mails</title>
		</head>
	<body style=font-family:arial align=center>
		<table>
			<tr><th>Zeitraum</th><th>Eingehend</th><th>Ausgehend</th></tr>
			<tr><td>Letzte 30 Tage</td><td>" + $inboundencrypted + " Mails verschlüsselt<br>" + $inboundnonencrypted + "Mails unverschlüsselt</td><td>" + $outboundencrypted + " Mails verschlüsselt<br>" + $outboundnonencrypted + "Mails unverschlüsselt</td></tr>"

$outboundencrypted = 0
$outboundnonencrypted = 0
$inboundencrypted = 0
$inboundnonencrypted = 0

for($x=0;$x -lt 30;$x++){

    $MessageTracks = Get-NspMessageTrack -From (Get-Date).AddDays(-31-$x) -To (Get-Date).AddDays(-30-$x)    
   
    Foreach ($MessageTrack in $MessageTracks)
    {
        if ($MessageTrack.WasReceivedFromRelayServer)
        {
            if ($MessageTrack.Encrypted -notlike 'None')
            {
                $outboundencrypted+=1
            }
            else
            {
                $outboundnonencrypted+=1
            }
        }
        else
        {
            if ($MessageTrack.Encrypted -notlike 'None')
            {
                $inboundencrypted+=1
            }
            else
            {
                $inboundnonencrypted+=1
            }
        }
    }
}
$global:htmlout+="<tr><td>Vorherige 30 Tage</td><td>" + $inboundencrypted + " Mails verschlüsselt<br>" + $inboundnonencrypted + "Mails unverschlüsselt</td><td>" + $outboundencrypted + " Mails verschlüsselt<br>" + $outboundnonencrypted + "Mails unverschlüsselt</td></tr>
        </table>"

Write-Host "Report Generated Successfully"

"Sending report to $ReportRecipient"

$message= New-Object Net.Mail.MailMessage
$smtp = New-Object Net.Mail.SmtpClient ($smtphost)
$message.From = $ReportSender
$message.To.Add($ReportRecipient)
$message.subject = "NoSpamProxy Verschlüsselungs-Report der letzten " + $Age + " Tage vom " + (Get-Date)
$message.IsBodyHtml = $true
$message.Body = $global:htmlout
$smtp.send($message)
$message.Dispose();