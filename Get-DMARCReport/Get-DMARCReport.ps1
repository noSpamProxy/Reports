Param
    (
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$false)][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	[Parameter(Mandatory=$false)][string] $ReportSubject = "DMARC Auswertung der letzten Woche",
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
$reportFileName = $Env:TEMP + "\dmarc-report.html"
$XMLreportFileName = $Env:TEMP + "\dmarc-report.xml"
$CSVreportFilename = $Env:TEMP + "\dmarc-report.csv"
$senderDomain = $null
$dkimsenderDomain = $null
$messageID = $null
$spfStatus = $null
$dkimStatus = $null
$dmarcStatus = $null
$wouldhavepassed = $null
$htmlbody2 = $null
$htmlout = $null
$elements = @()

Write-Host "Getting MessageTrackInformation.."
$messageTracks = Get-NSPMessageTrack -Status Success -From $dateStart -Directions FromExternal -WithOperations -WithActions -WithAddresses
Write-Host "Done."
Write-Host "Creating Report...."

foreach ($messageTrack in $messageTracks){
	$operations = $messageTrack.Operations
	foreach ($operation in $operations ){
		if ($operation.Operation.Type -eq "DmarcValidationEntry"){
			$dmarcCheckResult = $operation.Operation.Data
			$dmarchCheckDetails = $dmarcCheckResult |ConvertFrom-Json
			$addresses = $messageTrack.Addresses
			$senderDomain = ($addresses | ? {$_.AddressType -eq "Sender"} | select "Address").Address
			$messageID = $messageTrack.MessageID
			$emailcreated = $messageTrack.Sent.LocalDateTime
			$dmarceffectivePolicy = $dmarchCheckDetails.EffectivePolicy
			$dkimsenderDomain = $dmarchCheckDetails.OrganizationalDomain
			$dmarcsenderDomain = $dmarchCheckDetails.Rfc5322FromDomain
			$dmarcvalidationResult = $dmarchCheckDetails.ValidationResult
			$dkimResult = $dmarchCheckDetails.DkimResult
			$spfResult = $dmarchCheckDetails.SpfResult
			$dkimAlignment = $dmarchCheckDetails.DkimAlignment
			$spfAlignment = $dmarchCheckDetails.SpfAlignment
			$applicablePolicy = $dmarchCheckDetails.ApplicablePolicy
			$dmarcisvalid = $dmarchCheckDetails.validationResult
			foreach ($spfCheckEntry in $spfResult){
				if ($spfCheckEntry.Result -eq "Pass")
				{
				$spfResultDetail +=($spfCheckEntry.Domain +": OK")
				break
				}
				elseif ($spfCheckEntry.Result -ne "None")
				{
				$spfResultDetail +=($spfCheckEntry.Domain +": Failure<br>")
				}
				else
				{
				$spfResultDetail +=($spfCheckEntry.Domain +": None<br>")
				}
			}
			if ($dkimResult.Failures -eq "None"){
				$dkimResultDetail = "DKIM Signature present and OK"
			}
			else
			{
				$dkimResultDetail = "DKIM Signature not present or wrong"
			}
			$htmlbody2 +=("<tr><td><h3> "+$senderDomain +"</h3></td><td><h3> " +$dmarcsenderDomain +"</h3></td><td><h3>" +$spfAlignment +"</h3></td><td><h3>" +$spfResultDetail +"</h3></td><td><h3>"+$dkimAlignment +"</h3></td><td><h3>"+$dkimResultDetail +"</h3></td><td><h3>"+$dmarcvalidationResult +"</h3></td><td><h3>"+$applicablePolicy +"</h3></td><td><h3>"+$dmarceffectivePolicy +"</h3></td><td><h3> "+$emailcreated +"</h3></td><td><h3>"+$messageID +"</h3></td></tr>")
			$o = New-Object PsObject
			$o | Add-Member NoteProperty SenderDomain $senderDomain
			$o | Add-Member NoteProperty DMARCSenderDomain $dmarcsenderDomain
			$o | Add-Member NoteProperty SPFAlignment $spfAlignment
			$o | Add-Member NoteProperty SPFCheckResults $spfResultDetail
			$o | Add-Member NoteProperty DKIMAlignment $dkimAlignment
			$o | Add-Member NoteProperty DKIMCheckResults $dkimResultDetail
			$o | Add-Member NoteProperty DMARCValidationResult $dmarcvalidationResult
			$o | Add-Member NoteProperty ApplicablePolicy $applicablePolicy
			$o | Add-Member NoteProperty WhatIf $dmarceffectivePolicy
			
			$elements +=$o
			$spfResultDetail = $null
		}
	}
}

Write-Host "Done."
Write-Host "Generating and sending report"

$htmlbody1 ="<html>
			<head>
				<title>DMARC Auswertung der letzten Woche</title>
				<style>
	      			table, td, th { border: 1px solid black; border-collapse: collapse; white-space:nowrap; }
					#headerzeile         {background-color: #DDDDDD;}
	    		</style>
			</head>
		<body style=font-family:arial>
			<h1>DMARC Auswertung der letzten Woche</h1>
			<br>
			<table>
				<tr id=headerzeile>
					<td><h3>Sender-Domain</h3></td><td><h3>DMARC-Sender-Domain</h3></td><td><h3>SPF-Alignment</h3></td><td><h3>SPF-Check-Detail</h3></td><td><h3>DKIM-Alignment</h3></td><td><h3>DKIM-Check-Detail</h3></td><td><h3>DMARC Check</h3></td><td><h3>DMARC Policy</h3></td><td><h3>What if</h3></td><td><h3>Datum</h3></td><td><h3>Message ID</h3></td>
				</tr>
				"

$htmlbody3="</table>
		</body>
		</html>"

$htmlout=$htmlbody1+$htmlbody2+$htmlbody3
$htmlout | Out-File $reportFileName
(ConvertTo-Xml -notypeinformation $elements).OuterXml |out-file $XMLreportFileName
$elements.GetEnumerator() |Export-Csv -Path $CSVreportFilename -NoTypeInformation

Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $ReportRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der DMARC-Auswertung der letzten Woche." -Attachments ($reportFileName, $XMLreportFileName, $CSVreportFilename)

Write-Host "Done."
Write-Host "Doing some cleanup...."
Remove-Item $reportFileName
Remove-Item $XMLreportFileName
Remove-Item $CSVreportFilename
Write-Host "Done."