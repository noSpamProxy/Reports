param (
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$false)][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	[Parameter(Mandatory=$true)][string] $ReportRecipient,
	[Parameter(Mandatory=$false)][string] $ReportSubject = "Zertifikate in Quarantaene",
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

$certificatesInQuarantine = $null
$certificatesInQuarantineCount = $null
$reportFileName = $Env:TEMP + "\certificatereport.html"

Write-Host "Getting certificates in quarantine..."
$certificatesInQuarantine = get-nspcertificate -StoreIds Quarantined
if ($certificatesInQuarantine -eq $null){
	Write-Host "No certificates in quarantine!"
	exit
}

$certificatesInQuarantineCount = $certificatesInQuarantine.Count

foreach ($certificate in $certificatesInQuarantine){
	$certificateSubjectDN = $certificate.SubjectDistinguishedName
	$certificateIssuerDN = $certificate.IssuerDistinguishedName
	$certificateValidity = $certificate.ValidTo
	$htmlbody2 += "<tr><td width=150px>"+$certificateSubjectDN+"</td><td>"+$certificateIssuerDN+"</td><td>"+$certificateValidity+"</td></tr>"
}
$htmlbody1 ="<html>
			<head>
				<title>Zertifikate in Quarant$auml;ne</title>
				<style>
					table, td, th { border: 1px solid black; border-collapse: collapse; }
					#headerzeile         {background-color: #DDDDDD;}
				</style>
			</head>
		<body style=font-family:arial>
			<h1>"+$certificatesInQuarantineCount+" Zertifikate in Quarant&auml;ne</h1>
			<br>
			<table>
				<tr id=headerzeile>
					<td><h3>SubjectDistinguishedName</h3></td><td><h3>IssuerDistinguishedName</h3></td><td><h3>ValidTo</h3></td>
				</tr>
				"

$htmlbody3="</table>
		</body>
		</html>"
$htmlout=$htmlbody1+$htmlbody2+$htmlbody3
$htmlout | Out-File $reportFileName
Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $ReportRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der Auswertung der Zertifikate, die in Quarantaene liegen." -Attachments $reportFileName
Write-Host "Doing some cleanup.."
Remove-Item $reportFileName
Write-Host "Done."
