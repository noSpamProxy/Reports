Param
    (
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$false)][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung der abgelaufenen TLS-Zertifikate",
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

$dateStart = (Get-Date).AddDays(-$NumberOfDaysToReport)
$actualdate = Get-Date
$reportaddressesFileName = $Env:TEMP + "\expiredcertificates.html"
$recipientDomain = $null
$htmlbody2 = $null
$htmlout = $null
$elements = @()

Write-Host "Getting MessageTrackInformation.."
$messageTracks = Get-NSPMessageTrack -Status Success -From $dateStart -Directions FromLocal |Get-NSPMessageTrackdetails
Write-Host "Done."
Write-Host "Create Reportaddresses-File"

foreach ($messageTrack in $messageTracks)
	{
		$perRecipientStates = $messageTrack.Details.PerRecipientStatus
		foreach ($perRecipientStatus in $perRecipientStates)
			{
			$recipientDomain = $perRecipientStatus.RecipientAddress.Domain
			$connectionSecurityProtocol = $perRecipientStatus.ConnectionSecurity.SslProtocol
			if ($connectionSecurityProtocol -eq 0 ) {continue}
			$remoteIdentity = $perRecipientStatus.ConnectionSecurity.RemoteIdentity
			if ($remoteIdentity -eq $null) {continue}
			$TLSCertificateExpirationDate = $perRecipientStatus.ConnectionSecurity.RemoteIdentity.NotAfter
			$recipientIPAddress = $perRecipientStatus.TargetHost.Address.IPAddressToString
			$certificateName = $perRecipientStatus.ConnectionSecurity.RemoteIdentity.Subject
			if ($TLSCertificateExpirationDate -le $actualdate)
				{
					if($elements | where {$_.recipientDomain -eq $recipientDomain}) {continue}
					
					
					$o = New-Object PsObject
					$o | Add-Member NoteProperty recipientDomain $recipientDomain
					$o | Add-Member NoteProperty expirationDate $TLSCertificateExpirationDate
					$o | Add-Member NoteProperty IPAddress $recipientIPAddress
					$o | Add-Member NoteProperty CertificateSubject $certificateName
					
					$elements +=$o
					
				}
			}
	}

Write-Host "Done."
Write-Host "Generating and sending report"

$htmlbody1 ="<html>
			<head>
				<title>Auswertung der letzten Woche</title>
				<style>
	      			table, td, th { border: 1px solid black; border-collapse: collapse; white-space:nowrap; }
					#headerzeile         {background-color: #DDDDDD;}
	    		</style>
			</head>
		<body style=font-family:arial>
			<h1>Auflisting der E-Mail-Server mit abgelaufenen TLS-Zertifikaten der letzten "+$NumberOfDaysToReport+" Tage</h1>
			<br>
			<table>
				<tr id=headerzeile>
					<td><h3>Empf&auml;nger-Domain</h3></td><td><h3>Ablaufdatum</h3></td><td><h3>IP-Adresse</h3></td><td><h3>Zertifikatsinhaber</h3></td>
				</tr>
				"
$elements | sort recipientDomain | %{
	$htmlbody2 +=("<tr><td>"+$_.recipientDomain +"</td><td>" +$_.expirationDate +"</td><td>" +$_.IPAddress +"</td><td> " +$_.CertificateSubject +"</td></tr>")
}


$htmlbody3="</table>
		</body>
		</html>"

$htmlout=$htmlbody1+$htmlbody2+$htmlbody3
$htmlout | Out-File $reportaddressesFileName

Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $ReportRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie die Auflisting der E-Mail-Server mit abgelaufenen TLS-Zertifikaten." -Attachments $reportaddressesFileName

Write-Host "Done."
Write-Host "Doing some cleanup...."
Remove-Item $reportaddressesFileName
Write-Host "Done."