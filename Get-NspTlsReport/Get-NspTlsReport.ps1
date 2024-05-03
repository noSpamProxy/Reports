<#
.SYNOPSIS
  Name: Get-NspTlsReport.ps1
  Create report for in- and outbound E-Mails which did not use TLS.
  Requires NSP version 13.0.19147.917 or later.

.DESCRIPTION
  This script can be used to generate a report about E-Mails which where send or received without TLS
  It is possible to send the report via E-Mail and to filter the results for a specific time duration.
  This script only uses the NoSpamProxy Powershell Cmdlets to generate the report file.

.PARAMETER FromDate
  Mandatory if you like to use a timespan.
  Specifies the start date for the E-Mail filter.
  Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"  
  E.g.:
  	"2018-11-16 08:00" or "2018-11-16 20:00:00"

.PARAMETER NoTime
  Mandatory if you do not like to specify a time value in any kind of way.
  No value needs to be passed here <NoTime> is just a single switch.
  
.PARAMETER NspRule
  Specify a rule name which is defined in NSP as E-Mail filter.

.PARAMETER NumberOfDays
  Mandatory if you like to use a number of days for filtering.
  Specifies the number of days for which th E-Mails should be filtered.

.PARAMETER NumberOfHoursToReport
  Mandatory if you like to use a number of hours for filtering.
  Specifies the number of hours for which th E-Mails should be filtered.

.PARAMETER ReportFileName
  Default: TLSReport
	Sets the reports file name. No file extension required.
	
.PARAMETER ReportInterval
  Mandatory if you like to use a predifined timespan.
  Specifies a predifined timespan.
	Possible values are:
	daily, monthly, weekly
	The report will start at 00:00:00 o'clock and ends at 23:59:59 o'clock.
	The script call must be a day after the desired report end day.
  
.PARAMETER ReportRecipient
  Specifies the E-Mail recipient. It is possible to pass a comma seperated list to address multiple recipients. 
  E.g.: alice@example.com,bob@example.com

.PARAMETER ReportRecipientCSV
  Set a filepath to an CSV file containing a list of report E-Mail recipient. Be aware about the needed CSV format, please watch the provided example.

.PARAMETER ReportSender
  Default: NoSpamProxy Report Sender <nospamproxy@example.com>
  Sets the report E-Mail sender address.
  
.PARAMETER ReportSubject
  Default: NSP TLS Report
  Sets the report E-Mail subject.

.PARAMETER SmtpHost
  Specifies the SMTP host which should be used to send the report E-Mail.
  It is possible to use a FQDN or IP address.

.PARAMETER Status
  Specifies a filter to get only E-Mails which are matching the defined state.
  Possible values are: 
  None | Success | DispatcherError | TemporarilyBlocked | PermanentlyBlocked | PartialSuccess | DeliveryPending | Suppressed | DuplicateDrop | PutOnHold | All

.PARAMETER TenantPrimaryDomain
  Used to login into the desired NoSpamProxy tenant to run this script on.
  Only required if NoSpamProxy v14 is used in provider mode.

.PARAMETER ToDate
  Optional if you like to use a timespan.
  Specifies the end date for the E-Mail filter.
  Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"  
  E.g.:
  	"2018-11-16 08:00" or "2018-11-16 20:00:00"

.OUTPUTS
  Report is stored under %TEMP%\TLSReport.html unless a custom <ReportFileName> parameter is given.

.NOTES
  Version:        1.1.0
  Author:         Jan Jaeschke
  Creation Date:  2022-06-07
  Purpose/Change: added TLS version details
  
.LINK
  https://www.nospamproxy.de
  https://github.com/noSpamProxy

.EXAMPLE
  .\Get-NspTlsReport.ps1 -NoTime -Status "Success" -ReportFileName "Example-Report" -ReportRecipient alice@example.com -ReportSender "NoSpamProxy Report Sender <nospamproxy@example.com>" -ReportSubject "Example Report" -SmtpHost mail.example.com
  
.EXAMPLE
  .\Get-NspTlsReport.ps1 -FromDate: "2018-10-14 08:00" -ToDate: "2018-10-14 20:00" -NoMail
  It is mandatory to specify <FromDate>. Instead <ToDate> is optional.
  These parameters can be combined with all other parameters except <NumberOfDaysToReport> and <NumberOfHoursToRepor>.

.EXAMPLE 
  .\Get-NspTlsReport.ps1 -NumberOfDaysToReport 7 -NoMail
  You can combine <NumberOfDaysToReport> with all other parameters except <FromDate>, <ToDate> and <NumberOfHoursToRepor>.
  
.EXAMPLE 
  .\Get-NspTlsReport.ps1 -NumberOfHoursToReport 12 -NoMail
  You can combine <NumberOfHoursToReport> with all other parameters except <FromDate>, <ToDate> and <NumberOfDaysToReport>.
	
.EXAMPLE
	.\Get-NspTlsReport.ps1 -ReportInterval weekly -NoMail
	You can combine <ReportInterval> with all other parameters except <FromDate>, <ToDate>, <NumberOfDaysToReport> and <NumberOfHoursToReport>.
  
.EXAMPLE
  .\Get-NspTlsReport.ps1 -NoTime -NoMail -NspRule "All other inbound mails"
  
.EXAMPLE
  .\Get-NspTlsReport.ps1 -NoTime -SmtpHost mail.example.com -ReportRecipientCSV "C:\Users\example\Documents\email-report.csv"
  The CSV have to contain the header "Email" else the mail addresses cannot be read from the file.
  It is possible to combine <ReportRecipientCSV> with <ReportRecipient>.
  E.g: email-report.csv
  User,Email
  user1,user1@example.com
  user2,user2@example.com
#>
param (
# userParams are used for filtering
	# set start date for filtering
	[Parameter(Mandatory=$true, ParameterSetName="dateSpanSet")][ValidatePattern("^([0-9]{4})-?(1[0-2]|0[1-9])-?(3[01]|0[1-9]|[12][0-9])\s(2[0-3]|[01][0-9]):?([0-5][0-9]):?([0-5][0-9])?$")][string] $FromDate,
	# set end date for filtering
	[Parameter(Mandatory=$false, ParameterSetName="dateSpanSet")][ValidatePattern("^([0-9]{4})-?(1[0-2]|0[1-9])-?(3[01]|0[1-9]|[12][0-9])\s(2[0-3]|[01][0-9]):?([0-5][0-9]):?([0-5][0-9])?$")][string] $ToDate,
	# set number of days for filtering
	[Parameter(Mandatory=$true, ParameterSetName="numberOfDaysSet")][ValidatePattern("[0-9]+")][string] $NumberOfDaysToReport,
	# set number of hours for filtering
	[Parameter(Mandatory=$true, ParameterSetName="numberOfHoursSet")][int] $NumberOfHoursToReport,
	# set reporting intervall
	[Parameter(Mandatory=$true, ParameterSetName="reportIntervalSet")][ValidateSet('daily','weekly','monthly')][string] $ReportInterval,
	# if set not time duration have to be set the E-Mail of the last few hours will be filtered
	[Parameter(Mandatory=$true, ParameterSetName="noTimeSet")][switch] $NoTime, # no userParam just here for better Get-Help output
	# set E-Mail status which will be filtered
	[Parameter(Mandatory=$false)][string] $Status = "Success",
	# set NSP Rule for filtering
	[Parameter(Mandatory=$false)][string] $NspRule,
	# additional params are used for additional actions
	# generate the report but do not send an E-Mail
	[Parameter(Mandatory=$false)][switch]$NoMail, 
	# change report filename
	[Parameter(Mandatory=$false)][string] $ReportFileName = "TLSReport" ,
	# set report recipient only valid E-Mail addresses are allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+?<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*>?$")][string[]] $ReportRecipient,
	# set path to csv file containing report recipient E-Mail addresses
	[Parameter(Mandatory=$false)][string] $ReportRecipientCSV,
	# set report sender address only a valid E-Mail addresse is allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*>?$")][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	# change report E-Mail subject
	[Parameter(Mandatory=$false)][string] $ReportSubject = "NSP TLS Report",
	# set used SMTP host for sending report E-Mail only a valid  IP address or FQDN is allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^(((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))|(((?!-)[a-zA-Z0-9-]{0,62}[a-zA-Z0-9]\.)+[a-zA-Z]{2,63}))$")][string] $SmtpHost,
	# only needed for v14 with enabled provider mode
	[Parameter(Mandatory=$false)][string] $TenantPrimaryDomain	
)

#-------------------Functions----------------------
# process actual MessageTracks - compare Mail From and add htmlContent
function processMessageTracks($tmpMessageTracks, $messageDirection){
	$returnValues = @{}
	$returnValues.sendersWithoutTLS = @()
	foreach($messageTrack in $tmpMessageTracks){
		if($messageDirection -eq "FromExternal"){
			if ($useV14Queries) {
				$tls = ($messageTrack.TlsProtocol)
			} else {
				$tls = ($messageTrack.SenderConnectionSecurity)
			}
		}else{
			if ($useV14Queries) {
				$tls = ($messageTrack.DeliveryAttempts | Where-Object {$_.Status -eq "Success"}).TlsProtocol
			} else {
				$tls = ($messageTrack.DeliveryAttempts | Where-Object {$_.Status -eq "Success"}).ConnectionSecurity
			}
		}
		if($null -eq $tls -OR $tls -eq "None"){
			# number of the successfull send/received messages without 
			$returnValues.numberOfMessagesWithoutTls = $returnValues.numberOfMessagesWithoutTls+1
			if($messageDirection -eq "FromExternal"){
				$sender = ($messagetrack.Addresses|Where-Object{[string]$_.AddressType -eq "Sender"}).Address
			}else{
				$sender = ($messagetrack.Addresses|Where-Object{[string]$_.AddressType -eq "Recipient"}).Address
			}
			
			if($null -eq $sender){
				continue
			}else{
				$returnValues.sendersWithoutTLS += "$sender <br>"
			}
					} else{
			if($messageDirection -eq "FromExternal"){
				switch($tls){
					"Tls"{
						$returnValues.numberOfTls10 = $returnValues.numberOfTls10+1
						break;
					}
					"Tls11"{
						$returnValues.numberOfTls11 = $returnValues.numberOfTls11+1
						break;
					}
					"Tls12"{
						$returnValues.numberOfTls12 = $returnValues.numberOfTls12+1
						break;
					}
					"Tls13"{
						$returnValues.numberOfTls13 = $returnValues.numberOfTls13+1
						break;
					}
				}
			}
		}
	}
	$returnValues.sendersWithoutTLS = $returnValues.sendersWithoutTLS | Get-Unique
	return $returnValues
}
# get data from message tracks and call the processing
# input the desired direction of messages
function getData($messageDirection){
	$returnValues = @{}

	# number of messages which will be skipped by Get-NspMessageTrack, will increase by 100 at each call
	$skipMessageTracks = 0
	while($getMessageTracks -eq $true){	
		if($skipMessageTracks -eq 0){
			$messageTracks = Get-NSPMessageTrack @cleanedParams -WithOperations -WithDeliveryAttempts -WithTlsSettings -WithActions -WithAddresses -Directions $messageDirection -First 100
			# number of messages which where send/received successfully
			$returnValues.numberOfMessages += $messageTracks.Count
		}else{
			$messageTracks = Get-NSPMessageTrack @cleanedParams -WithOperations -WithDeliveryAttempts -WithTlsSettings -WithActions -WithAddresses -Directions $messageDirection -First 100 -Skip $skipMessageTracks
			# number of messages which where send/received successfully
			$returnValues.numberOfMessages += $messageTracks.Count
		}
		$processResult = processMessageTracks $messageTracks $messageDirection
		$returnValues.numberOfMessagesWithoutTls += $processResult.numberOfMessagesWithoutTls
		$returnValues.sendersWithoutTLS += $processResult.sendersWithoutTLS
		$returnValues.numberOfTls10 += $processResult.numberOfTls10
		$returnValues.numberOfTls11 += $processResult.numberOfTls11
		$returnValues.numberOfTls12 += $processResult.numberOfTls12
		$returnValues.numberOfTls12 += $processResult.numberOfTls13
		# exit condition
		if($messageTracks){
			$skipMessageTracks = $skipMessageTracks+100
			Write-Host $skipMessageTracks
		}else{
			$getMessageTracks = $false
			break
		}
	}
	return $returnValues
}
# generate HMTL report
function createHTML($resultFromExternal, $percentFromExternal, $resultFromLocal, $percentFromLocal) {
	$percentFromExternalTls10 = [Math]::Round($($resultFromExternal.numberOfTls10)/$($resultFromExternal.numberOfMessages)*100,2)
	$percentFromExternalTls11 = [Math]::Round($($resultFromExternal.numberOfTls11)/$($resultFromExternal.numberOfMessages)*100,2)
	$percentFromExternalTls12 = [Math]::Round($($resultFromExternal.numberOfTls12)/$($resultFromExternal.numberOfMessages)*100,2)
	$percentFromExternalTls13 = [Math]::Round($($resultFromExternal.numberOfTls13)/$($resultFromExternal.numberOfMessages)*100,2)
	$htmlOut =
"<html>
	<head>
		<title>Auswertung der TLS-Sicherheit</title>
		<style>
			table, td, th { border: 1px solid black; border-collapse: collapse; padding:10px; text-align:center;}
			#headerzeile         {background-color: #DDDDDD;}
		</style>
	</head>
	<body style=font-family:arial>
		<h1>Auswertung der TLS-Sicherheit</h1>
		<br>
		Anzahl aller eingehenden Verbindungen: $($resultFromExternal.numberOfMessages)<br>
		Davon waren unverschl&uuml;sselt: $($resultFromExternal.numberOfMessagesWithoutTls) ($percentFromExternal%)<br><br>
		Davon waren Tls 1.0 verschl&uuml;sselt: $($resultFromExternal.numberOfTls10) ($percentFromExternalTls10%)<br>
		Davon waren Tls 1.1 verschl&uuml;sselt: $($resultFromExternal.numberOfTls11) ($percentFromExternalTls11%)<br>
		Davon waren Tls 1.2 verschl&uuml;sselt: $($resultFromExternal.numberOfTls12) ($percentFromExternalTls12%)<br>
		Davon waren Tls 1.3 verschl&uuml;sselt: $($resultFromExternal.numberOfTls13) ($percentFromExternalTls13%)<br><br>
		Anzahl aller ausgehenden Verbindungen: $($resultFromLocal.numberOfMessages)<br>
		Davon waren unverschl&uuml;sselt: $($resultFromLocal.numberOfMessagesWithoutTls) ($percentFromLocal%)<br><br>
		Adressen, die unverschl&uuml;sselt geschickt haben: <br>
		$($($resultFromExternal.sendersWithoutTLS) | Out-String)
		<br><br>
		Adressen, an die unverschl&uuml;sselt geschickt wurde: <br>
		$($($resultFromLocal.sendersWithoutTLS) | Out-String)
	</body>
</html>"
	$htmlOut | Out-File "$reportFile"
}
# send report E-Mail 
function sendMail($ReportRecipient, $ReportRecipientCSV){ 
	if ($ReportRecipient -and $ReportRecipientCSV){
		$recipientCSV = Import-Csv $ReportRecipientCSV
		$mailRecipient = @($ReportRecipient;$recipientCSV.Email)
	}
	elseif($ReportRecipient){
		$mailRecipient = $ReportRecipient
	}
	elseif($ReportRecipientCSV){
		$csv = Import-Csv $ReportRecipientCSV
		$mailRecipient = $csv.Email
	}
	if ($SmtpHost -and $mailRecipient){
		Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $mailRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie einen automatisch genrerierten Bericht vom NoSpamProxy" -Attachments $reportFile
	}
}

#-------------------Variables----------------------
$fileDate = Get-Date -UFormat "%Y-%m-%d"
if ($NumberOfHoursToReport){
	$FromDate = (Get-Date).AddHours(-$NumberOfHoursToReport)
}
if ($ReportInterval){
	# equals the day where the script runs
	$reportEndDay = (Get-Date -Date ((Get-Date).AddDays(-1)) -UFormat "%Y-%m-%d")
	switch ($ReportInterval){
		'daily'{
			$reportStartDay = $reportEndDay
		}
		'weekly'{
			$reportStartDay = (Get-Date -Date ((Get-Date).AddDays(-7)) -UFormat "%Y-%m-%d")
		}
		'monthly'{
			$reportStartDay = (Get-Date -Date ((Get-Date).AddMonths((-1))) -UFormat "%Y-%m-%d")
		}
	}
	$FromDate = "$reportStartDay 00:00:00"
	$ToDate = "$reportEndDay 23:59:59"
	$fileDate = $reportEndDay
  $reportFile =  "$ENV:TEMP\" + "$fileDate-$ReportFileName" + ".html"
}
$reportFile =  "$ENV:TEMP\" + "$ReportFileName" + ".html"

if ($NumberOfHoursToReport){
	$FromDate = (Get-Date).AddHours(-$NumberOfHoursToReport)
}
# create hashtable which will preserve the order of the added items and mapps userParams into needed parameter format
$userParams = [ordered]@{ 
From = $FromDate
To = $ToDate
Age = $NumberOfDaysToReport
Rule = $NspRule
Status = $Status
} 
# for loop problem because hashtable have no indices to access items, this is a workaround
# new hashtable which only holds non empty userParams
$cleanedParams=@{}
# this loop removes all empty userParams and add the otherones into the new hashtable
foreach ($userParam in $userParams.Keys) {
	if ($($userParams.Item($userParam)) -ne "") {
		$cleanedParams.Add($userParam, $userParams.Item($userParam))
	}
}
# end workaround
# condition to run Main part, if false program will end
$getMessageTracks = $true


#--------------------Main-----------------------
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
if ($nspVersion -ge '14.0.231') {
	$useV14Queries = $true
}
# check NSP version for compatibility 
if(!((Get-NspIntranetRole).Version -ge [version]"13.0.19147.917")){
	throw "A minimum of NSP Version 13.0.19147.917 is needed, please update. Your version: $((Get-NspIntranetRole).Version)"
}

Write-Host "Getting data for E-Mails from External"
$resultFromExternal = getData "FromExternal"
if($null -eq $($resultFromExternal.sendersWithoutTLS)){
	$resultFromExternal.sendersWithoutTLS = 0
	$percentFromExternal = 0
}else{
	$percentFromExternal = [Math]::Round($($resultFromExternal.numberOfMessagesWithoutTls)/$($resultFromExternal.numberOfMessages)*100,2)
}

Write-Host "Getting data for E-Mails from local"
$resultFromLocal = getdata "FromLocal"
if($null -eq $($resultFromLocal.numberOfMessagesWithoutTls)){
	$resultFromLocal.numberOfMessagesWithoutTls = 0
	$percentFromLocal = 0
}else{
	$percentFromLocal = [Math]::Round($($resultFromLocal.numberOfMessagesWithoutTls)/$($resultFromLocal.numberOfMessages)*100,2)
}

Write-Host "Generating Report"
createHTML $resultFromExternal $percentFromExternal $resultFromLocal $percentFromLocal
# send mail if <NoMail> switch is not used and delete temp report file
if (!$NoMail){
	sendMail $ReportRecipient $ReportRecipientCSV
	Remove-Item $reportFile
} else {
	Move-Item -Path $reportFile -Destination "$PSScriptRoot\$(Split-Path -Path $reportFile -Leaf)"
}
Write-Host "Skript durchgelaufen"