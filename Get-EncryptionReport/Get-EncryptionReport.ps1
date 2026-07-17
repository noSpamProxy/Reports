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
  [Parameter(Mandatory=$true, ParameterSetName="default")][string] $smtphost,
  [Parameter(Mandatory=$true, ParameterSetName="default")][string] $ReportSender,
  [Parameter(Mandatory=$true, ParameterSetName="default")][string] $ReportRecipient,
  # only needed for v14 with enabled provider mode
	[Parameter(Mandatory=$false)][string] $TenantPrimaryDomain,
  # Optional credential. When supplied, BASIC authentication is used. Omit it to use integrated (NTLM) auth, or force NTLM with -UseIntegratedAuth.
  [Parameter(Mandatory=$false)][PSCredential] $Credential,
  [Parameter(Mandatory=$false)][string] $CredentialFileLocation = "$PSScriptRoot\NspBasicAuthCredential.bin",
  [Parameter(Mandatory=$true, ParameterSetName="setCredentialSet")][switch] $SetCredential,
  [Parameter(Mandatory=$false)][switch] $UseIntegratedAuth
)

if (-not [System.IO.Path]::IsPathRooted($CredentialFileLocation)) {
  $CredentialFileLocation = Join-Path -Path $PSScriptRoot -ChildPath $CredentialFileLocation
}

# create an encrypted local credential file using DPAPI LocalMachine
function Set-loginPass {
  Add-Type -AssemblyName System.Security
  $username = Read-Host -Prompt 'Input your username'
  $securePass = Read-Host -Prompt 'Input your user password' -AsSecureString
  $plainPass = [System.Net.NetworkCredential]::new('', $securePass).Password
  $inBytes = [System.Text.Encoding]::Unicode.GetBytes($plainPass)
  $protected = [System.Security.Cryptography.ProtectedData]::Protect($inBytes, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
  $payload = @{ FormatVersion = 1; Username = $username; Password = [System.Convert]::ToBase64String($protected) } | ConvertTo-Json -Compress
  [System.IO.File]::WriteAllText($CredentialFileLocation, $payload, [System.Text.Encoding]::UTF8)
}

# read credential file and return a PSCredential object if possible
function Get-loginPass {
  Add-Type -AssemblyName System.Security
  if (-not (Test-Path $CredentialFileLocation)) {
    return $null
  }

  try {
    $content = [System.IO.File]::ReadAllText($CredentialFileLocation, [System.Text.Encoding]::UTF8)
    $payload = $content | ConvertFrom-Json
    if ([string]::IsNullOrWhiteSpace($payload.Username) -or [string]::IsNullOrWhiteSpace($payload.Password)) {
      Write-Warning "Credential file '$CredentialFileLocation' does not use the expected username/password JSON format. Recreate it with -SetCredential. Falling back to integrated authentication."
      return $null
    }

    $protected = [System.Convert]::FromBase64String($payload.Password)
    $rawKey = [System.Security.Cryptography.ProtectedData]::Unprotect($protected, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    $plainPass = [System.Text.Encoding]::Unicode.GetString($rawKey)
    $securePassword = ConvertTo-SecureString -String $plainPass -AsPlainText -Force
    return New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $payload.Username, $securePassword
  }
  catch {
    Write-Warning "Credential file '$CredentialFileLocation' cannot be loaded: $($_.Exception.Message) Recreate it with -SetCredential if it was created on another machine or with an older password-only format. Falling back to integrated authentication."
    return $null
  }
}

if($SetCredential){
  Set-loginPass
  EXIT
}

$resolvedCredential = $null
if (-not $UseIntegratedAuth) {
  if ($Credential) {
    $resolvedCredential = $Credential
  }
  else {
    $resolvedCredential = Get-loginPass
  }
}

$nspVersion = (Get-ItemProperty -Path HKLM:\SOFTWARE\NoSpamProxy\Components -ErrorAction SilentlyContinue).'Intranet Role'
if ($nspVersion -gt '14.0') {
	try {
		$connectParams = @{ IgnoreServerCertificateErrors = $true; ErrorAction = 'Stop' }; if ($resolvedCredential) { $connectParams['Credential'] = $resolvedCredential.GetNetworkCredential() }
		Connect-Nsp @connectParams
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
      $connectParams = @{ IgnoreServerCertificateErrors = $true; PrimaryDomain = $TenantPrimaryDomain }; if ($resolvedCredential) { $connectParams['Credential'] = $resolvedCredential.GetNetworkCredential() }
			Connect-Nsp @connectParams
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