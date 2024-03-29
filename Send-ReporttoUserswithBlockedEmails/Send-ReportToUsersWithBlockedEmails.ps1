Param
(
[Parameter(Mandatory=$true)][string] $SMTPHost,
[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 1,
[Parameter(Mandatory=$false)][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung der abgewiesenen E-Mails an Sie",
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
$reportaddressesFileName = [System.IO.Path]::Combine($Env:TEMP, "reportaddresses.txt")

"Getting MessageTrackInformation.."
$messageTracks = Get-NSPMessageTrack -Status PermanentlyBlocked -From $dateStart -Directions FromExternal -WithAddresses
"Create Reportaddresses-File"

$entries = @{}
foreach ($messageTrack in $messageTracks){
	$addresses = $messageTrack.Addresses
	foreach ($addressEntry in $addresses){
		if ($addressEntry.AddressType -eq "Recipient"){
			$messageRecipient = $addressEntry.Address
			$list = $entries[$messageRecipient]
			if (!$list) {
				$list = @($messagetrack)
			}
			else
			{
				$list += $messageTrack
			}
			$entries[$messageRecipient] = $list
			}
	}
}
Set-Content $reportaddressesFileName $existingAddresses

"Generating and sending reports for the following e-mail addresses:"

$entries.GetEnumerator() | ForEach-Object {

	$_.Name
	$reportFileName = $Env:TEMP + "\reject-analysis.html"

	$htmlbody1 ="<html>
			<head>
				<title>Abgewiesene E-Mails an Sie</title>
				<style>
					table, td, th { border: 1px solid black; border-collapse: collapse; }
					#headerzeile         {background-color: #DDDDDD;}
				</style>
			</head>
		<body style=font-family:arial>
			<h1>Abgewiesene E-Mails an Sie</h1>
			<br>
			<table>
				<tr id=headerzeile>
					<td><h3>Uhrzeit</h3></td><td><h3>Absender</h3></td><td><h3>Betreff</h3></td>
				</tr>
				"

	$htmlbody2 =""
	foreach ($validationItem in $_.Value) 
	{
		$NSPStartTime = $validationItem.Sent.LocalDateTime
		$addresses2 = $validationItem.Addresses
		$NSPSender = ($addresses2 | ? {$_.AddressType -eq "Sender"} | select "Address").Address		
		$NSPSubject = $validationItem.Subject
		$htmlbody2 += "<tr><td width=150px>$NSPStartTime</td><td>$NSPSender</td><td>$NSPSubject</td></tr>"
	}

	$htmlbody3="</table>
		</body>
		</html>"
	$htmlout=$htmlbody1+$htmlbody2+$htmlbody3

	$htmlout | Out-File $reportFileName
	Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $_.Name -Subject $ReportSubject -BodyAsHtml -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der Auswertung der abgewiesenen E-Mails aufgrund von Anh&auml;ngen an der E-Mail." -Attachments $reportFileName
	Remove-Item $reportFileName
}
"Doing some cleanup...."
Remove-Item $reportaddressesFileName

