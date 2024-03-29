<#
.SYNOPSIS
  Name: Get-NspSandboxStatistic.ps1
  Create a Cyren Sandbox statistic for rejected emails.

.DESCRIPTION
  This script can be used to generate a Cyren Sandbox statistic for rejected emails.
  It is possible to send the report via E-Mail and to filter the results for a specific time duration.
  This script only uses the NoSpamProxy Powershell Cmdlets to generate the report file.

.PARAMETER FromDate
  Mandatory if you like to use a timespan.
  Specifies the start date for the E-Mail filter.
  Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"  
  E.g.:
  	"2018-11-16 08:00" or "2018-11-16 20:00:00"

.PARAMETER ToDate
  Optional if you like to use a timespan.
  Specifies the end date for the E-Mail filter.
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
  Default: NSP Sandbox statistic
  Sets the report E-Mail subject.

.PARAMETER SmtpHost
  Specifies the SMTP host which should be used to send the report E-Mail.
  It is possible to use a FQDN or IP address.

.PARAMETER TenantPrimaryDomain
  Used to login into the desired NoSpamProxy tenant to run this script on.
  Only required if NoSpamProxy v14 is used in provider mode.

.OUTPUTS
  none

.NOTES
  Version:        1.0.3
  Author:         Jan Jaeschke
  Creation Date:  2022-06-07
  Purpose/Change: added v14 support
  
.LINK
  https://www.nospamproxy.de
  https://github.com/noSpamProxy

.EXAMPLE
  .\Get-NspSandboxStatistic.ps1 -NoTime -ReportRecipient alice@example.com -ReportSender "NoSpamProxy Report Sender <nospamproxy@example.com>" -ReportSubject "Example Report" -SmtpHost mail.example.com
  
.EXAMPLE
  .\Get-NspSandboxStatistic.ps1 -FromDate: "2018-10-14 08:00" -ToDate: "2018-10-14 20:00:00" -NoMail
  It is mandatory to specify <FromDate>. Instead <ToDate> is optional.
  These parameters can be combined with all other parameters except <NumberOfDaysToReport> and <NumberOfHoursToRepor>.

.EXAMPLE 
  .\Get-NspSandboxStatistic.ps1 -NumberOfDaysToReport 7 -NoMail
  You can combine <NumberOfDaysToReport> with all other parameters except <FromDate>, <ToDate> and <NumberOfHoursToRepor>.
  
.EXAMPLE 
  .\Get-NspSandboxStatistic.ps1 -NumberOfHoursToReport 12 -NoMail
  You can combine <NumberOfHoursToReport> with all other parameters except <FromDate>, <ToDate> and <NumberOfDaysToReport>.
	
.EXAMPLE
	.\Get-NspSandboxStatistic.ps1 -ReportInterval weekly -NoMail
	You can combine <ReportInterval> with all other parameters except <FromDate>, <ToDate>, <NumberOfDaysToReport> and <NumberOfHoursToReport>.
  
.EXAMPLE
  .\Get-NspSandboxStatistic.ps1 -NoTime -NoMail -NspRule "All other inbound mails"
  
.EXAMPLE
  .\Get-NspSandboxStatistic.ps1 -NoTime -SmtpHost mail.example.com -ReportRecipientCSV "C:\Users\example\Documents\email-report.csv"
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
	# set NSP Rule for filtering
	[Parameter(Mandatory=$false)][string] $NspRule,
	# generate the report but do not send an E-Mail
	[Parameter(Mandatory=$false)][switch]$NoMail, 
	# set report recipient only valid E-Mail addresses are allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+?<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*>?$")][string[]] $ReportRecipient,
	# set path to csv file containing report recipient E-Mail addresses
	[Parameter(Mandatory=$false)][string] $ReportRecipientCSV,
	# set report sender address only a valid E-Mail addresse is allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*>?$")][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	# change report E-Mail subject
	[Parameter(Mandatory=$false)][string] $ReportSubject = "NSP Sandbox statistic",
	# set used SMTP host for sending report E-Mail only a valid  IP address or FQDN is allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^(((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))|((([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])))$")][string] $SmtpHost,
	# only needed for v14 with enabled provider mode
	[Parameter(Mandatory=$false)][string] $TenantPrimaryDomain
)

#-------------------Functions----------------------
# show wait animation while run $command specified script
function getData($command){
	$scroll = "/-\|/-\|"
	$idx = 0
	
	$job = Invoke-Command -ComputerName $env:ComputerName -ScriptBlock $command -AsJob
	
	$origpos = $host.UI.RawUI.CursorPosition
	$origpos.Y += 1
	
	while (($job.State -eq "Running") -and ($job.State -ne "NotStarted"))
	{
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host Please wait $scroll[$idx] -NoNewline
		$idx++
		if ($idx -ge $scroll.Length)
		{
			$idx = 0
		}
		Start-Sleep -Milliseconds 100
	}
	
	# It's over - clear the activity indicator.
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host ' '
	return Receive-Job $job
}
# send report E-Mail 
function sendMail($ReportRecipient, $ReportRecipientCSV, $mailBody){ 
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
		Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $mailRecipient -Subject $ReportSubject -Body $mailBody
	}
}

#-------------------Variables----------------------
$Status = "PermanentlyBlocked"
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
}
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

$command = {
	# needed to access cleanedParams - look for a better solution: https://adamtheautomator.com/invoke-command-remote/
	$customParams = $using:cleanedParams
	# condition to run Main part, if false program will end
	$getMessageTracks = $true
	# number of messages whcih will be skipped by Get-NspMessageTrack, will increase by 100 at each call
	$skipMessageTracks = 0

	$returnValues = @{msgCount = 0; sandboxCount = 0}
	#--------------------Main-----------------------
	while ($getMessageTracks -eq $true) {	
		if ($skipMessageTracks -eq 0) {
			$messageTracks = Get-NSPMessageTrack @customParams -WithOperations -Directions FromExternal -First 100
		}
		else {
			$messageTracks = Get-NSPMessageTrack @customParams -WithOperations -Directions FromExternal -First 100 -Skip $skipMessageTracks
		}
		$returnValues.msgCount = $returnValues.msgCount + $messageTracks.Count
		$sandBoxMsg = ($messageTracks.Operations.Operation | Where-Object { $_.Type -like "SandboxAnalysis" })
		if ($null -ne $sandBoxMsg.Data) {
			$returnValues.sandboxCount += (($sandBoxMsg.Data | ConvertFrom-Json).sandboxAnalysisResults.Status -like "Infected").Count
		}
		# exit condition
		if ($messageTracks) {
			$skipMessageTracks = $skipMessageTracks + 100
			Write-Verbose $skipMessageTracks
		}
		else {
			$getMessageTracks = $false
			break
		}
	}
	return $returnValues
}
$getDataReturnValues = getData $command
if($getDataReturnValues.msgCount -ne 0 -OR $getDataReturnValues.sandboxCount -ne 0){
	$sandboxPercent = [Math]::Round($getDataReturnValues.sandboxCount/$getDataReturnValues.msgCount*100,2)
}else{
	$sandboxPercent = "0"
}
$mailBody = @"
Number of blocked messages: $($getDataReturnValues.msgCount)
Caused by Cyren sandbox: $($getDataReturnValues.sandboxCount) ($sandboxPercent%)
"@
Write-Output $mailBody
# send mail if <NoMail> switch is not used and delete temp report file
if (!$NoMail){
	sendMail $ReportRecipient $ReportRecipientCSV $mailBody
}
Write-Host "Skript durchgelaufen"