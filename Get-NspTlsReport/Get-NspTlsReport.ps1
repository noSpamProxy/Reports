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

.PARAMETER ToDate
  Optional if you like to use a timespan.
  Specifies the end date for the E-Mail filter.
  Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"  
  E.g.:
  	"2018-11-16 08:00" or "2018-11-16 20:00:00"

.OUTPUTS
  Report is stored under %TEMP%\TLSReport.html unless a custom <ReportFileName> parameter is given.

.NOTES
  Version:        1.0.3
  Author:         Jan Jaeschke
  Creation Date:  2020-01-02
  Purpose/Change: added validation pattern for FromDate and ToDate parameters
  
.LINK
  https://https://www.nospamproxy.de
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
	[Parameter(Mandatory=$false)][ValidatePattern("^(((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))|(((?!-)[a-zA-Z0-9-]{0,62}[a-zA-Z0-9]\.)+[a-zA-Z]{2,63}))$")][string] $SmtpHost
)

#-------------------Functions----------------------
# process actual MessageTracks - compare Mail From and add htmlContent
function processMessageTracks($tmpMessageTracks, $messageDirection){
	$returnValues = @{}
	$returnValues.sendersWithoutTLS = @()
	foreach($messageTrack in $tmpMessageTracks){
		if($messageDirection -eq "FromExternal"){
			$tls = ($messageTrack.SenderConnectionSecurity)
		}else{
			$tls = ($messageTrack.DeliveryAttempts | Where-Object {$_.Status -eq "Success"}).ConnectionSecurity
		}
		if($null -eq $tls){
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
}
Write-Host "Skript durchgelaufen"
# SIG # Begin signature block
# MIIbigYJKoZIhvcNAQcCoIIbezCCG3cCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUa4SuvmAskUe73K3IkQbsgVec
# mmugghbWMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
# BQAwTDEgMB4GA1UECxMXR2xvYmFsU2lnbiBSb290IENBIC0gUjMxEzARBgNVBAoT
# Ckdsb2JhbFNpZ24xEzARBgNVBAMTCkdsb2JhbFNpZ24wHhcNMTYwNjE1MDAwMDAw
# WhcNMjQwNjE1MDAwMDAwWjBaMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFs
# U2lnbiBudi1zYTEwMC4GA1UEAxMnR2xvYmFsU2lnbiBDb2RlU2lnbmluZyBDQSAt
# IFNIQTI1NiAtIEczMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAjYVV
# I6kfU6/J7TbCKbVu2PlC9SGLh/BDoS/AP5fjGEfUlk6Iq8Zj6bZJFYXx2Zt7G/3Y
# SsxtToZAF817ukcotdYUQAyG7h5LM/MsVe4hjNq2wf6wTjquUZ+lFOMQ5pPK+vld
# sZCH7/g1LfyiXCbuexWLH9nDoZc1QbMw/XITrZGXOs5ynQYKdTwfmOPLGC+MnwhK
# kQrZ2TXZg5J2Yl7fg67k1gFOzPM8cGFYNx8U42qgr2v02dJsLBkwXaBvUt/RnMng
# Ddl1EWWW2UO0p5A5rkccVMuxlW4l3o7xEhzw127nFE2zGmXWhEpX7gSvYjjFEJtD
# jlK4PrauniyX/4507wIDAQABo4IBZDCCAWAwDgYDVR0PAQH/BAQDAgEGMB0GA1Ud
# JQQWMBQGCCsGAQUFBwMDBggrBgEFBQcDCTASBgNVHRMBAf8ECDAGAQH/AgEAMB0G
# A1UdDgQWBBQPOueslJF0LZYCc4OtnC5JPxmqVDAfBgNVHSMEGDAWgBSP8Et/qC5F
# JK5NUPpjmove4t0bvDA+BggrBgEFBQcBAQQyMDAwLgYIKwYBBQUHMAGGImh0dHA6
# Ly9vY3NwMi5nbG9iYWxzaWduLmNvbS9yb290cjMwNgYDVR0fBC8wLTAroCmgJ4Yl
# aHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9yb290LXIzLmNybDBjBgNVHSAEXDBa
# MAsGCSsGAQQBoDIBMjAIBgZngQwBBAEwQQYJKwYBBAGgMgFfMDQwMgYIKwYBBQUH
# AgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3JlcG9zaXRvcnkvMA0GCSqG
# SIb3DQEBCwUAA4IBAQAVhCgM7aHDGYLbYydB18xjfda8zzabz9JdTAKLWBoWCHqx
# mJl/2DOKXJ5iCprqkMLFYwQL6IdYBgAHglnDqJQy2eAUTaDVI+DH3brwaeJKRWUt
# TUmQeGYyDrBowLCIsI7tXAb4XBBIPyNzujtThFKAzfCzFcgRCosFeEZZCNS+t/9L
# 9ZxqTJx2ohGFRYzUN+5Q3eEzNKmhHzoL8VZEim+zM9CxjtEMYAfuMsLwJG+/r/uB
# AXZnxKPo4KvcM1Uo42dHPOtqpN+U6fSmwIHRUphRptYCtzzqSu/QumXSN4NTS35n
# fIxA9gccsK8EBtz4bEaIcpzrTp3DsLlUo7lOl8oUMIIE+zCCA+OgAwIBAgIMXyow
# wDWeCuKMV1r4MA0GCSqGSIb3DQEBCwUAMFoxCzAJBgNVBAYTAkJFMRkwFwYDVQQK
# ExBHbG9iYWxTaWduIG52LXNhMTAwLgYDVQQDEydHbG9iYWxTaWduIENvZGVTaWdu
# aW5nIENBIC0gU0hBMjU2IC0gRzMwHhcNMTkwNjA2MTM0NzIxWhcNMjIwNjA2MTM0
# NzIxWjB1MQswCQYDVQQGEwJERTEcMBoGA1UECBMTTm9yZHJoZWluLVdlc3RmYWxl
# bjESMBAGA1UEBxMJUGFkZXJib3JuMRkwFwYDVQQKExBOZXQgYXQgV29yayBHbWJI
# MRkwFwYDVQQDExBOZXQgYXQgV29yayBHbWJIMIIBIjANBgkqhkiG9w0BAQEFAAOC
# AQ8AMIIBCgKCAQEAstEMIFLNGUWS4uipXK3J6jJRBtI8+WjlUNal/WmOU4vSeJBC
# 4BkG8AsTZPd4KNEIVlbXi4MNV2eMtoQgyhRF1iQFGFXhqO0qxhYLArfUSEPPekL+
# t/ySEPVEurliH6Di1qfaFxceM+dXWG6ybrlOOZkHqow1PqBPfOUC54Rcyq6Co+mu
# qNvznCBPZSK4wvbiHCYb2pN0tnl7swP1q/K0ODB23wJathgKmLemW6Coz7L/sBHH
# vpgU1fVwi8huavjtQMFv0IRXiKZuHDnAugyNrEpJpFpQpxXLUpEN9Bn0GzmTth0N
# tVCMXVPeChj3qjvJEYP3GnGpY7K6O0Zc6Ao/jQIDAQABo4IBpDCCAaAwDgYDVR0P
# AQH/BAQDAgeAMIGUBggrBgEFBQcBAQSBhzCBhDBIBggrBgEFBQcwAoY8aHR0cDov
# L3NlY3VyZS5nbG9iYWxzaWduLmNvbS9jYWNlcnQvZ3Njb2Rlc2lnbnNoYTJnM29j
# c3AuY3J0MDgGCCsGAQUFBzABhixodHRwOi8vb2NzcDIuZ2xvYmFsc2lnbi5jb20v
# Z3Njb2Rlc2lnbnNoYTJnMzBWBgNVHSAETzBNMEEGCSsGAQQBoDIBMjA0MDIGCCsG
# AQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzAI
# BgZngQwBBAEwCQYDVR0TBAIwADA/BgNVHR8EODA2MDSgMqAwhi5odHRwOi8vY3Js
# Lmdsb2JhbHNpZ24uY29tL2dzY29kZXNpZ25zaGEyZzMuY3JsMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMB8GA1UdIwQYMBaAFA8656yUkXQtlgJzg62cLkk/GapUMB0GA1Ud
# DgQWBBQfMkfbwLvXrRRwHDgqny8W9JqFizANBgkqhkiG9w0BAQsFAAOCAQEATV/B
# SwkQkEbtB4JVCZBEowPzU2FdJzxS3LKg6NW2GX9vd3iHU/703AL8dqBSdoO6CREw
# /GV3pXtQhWDv1HVuCCRNk+rf4NooDMgxtNZFaAcKn8Zto+/a+4f01URf1LObbIeg
# bHByaBzlLv1FW3v/ilsLCs+KJ8Vkp/qG1gxac/KR79yLTXa1wgNkIvAtCz9LRlqf
# 0qUWubVC6Hg1s2EnuSs2d+v497zZRIp+UxkqLp3Uuvacp8VTl+NY3q064Fm2QyG5
# xwX8FWO+hwEF6mH2vh71icxXsRVADCgiOBX7S0l0M+zTVwnadPE6VlmLlcRo2Uv/
# /xrNfYi4zYch/b/ZtjCCBmowggVSoAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJ
# KoZIhvcNAQEFBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IElu
# YzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQg
# QXNzdXJlZCBJRCBDQS0xMB4XDTE0MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFow
# RzELMAkGA1UEBhMCVVMxETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdp
# Q2VydCBUaW1lc3RhbXAgUmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8A
# MIIBCgKCAQEAo2Rd/Hyz4II14OD2xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe
# 1VpjWwJJUNmDzm9m7t3LhelfpfnUh3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRk
# tlC39RKzc5YKZ6O+YZ+u8/0SeHUOplsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95O
# X+Koti1ZAmGIYXIYaLm4fO7m5zQvMXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P
# 7IwMNlt6wXq4eMfJBi5GEMiN6ARg27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZ
# VjCT88lhzNAIzGvsYkKRrALA76TwiRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/
# BAQDAgeAMAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/
# BgNVHSAEggG2MIIBsjCCAaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFS
# AEEAbgB5ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBj
# AGEAdABlACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBu
# AGMAZQAgAG8AZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQ
# AFMAIABhAG4AZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAg
# AEEAZwByAGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABp
# AGEAYgBpAGwAaQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwBy
# AGEAdABlAGQAIABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBl
# AC4wCwYJYIZIAYb9bAMVMB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLN
# MB0GA1UdDgQWBBRhWk0ktkkynUoqeRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0
# hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0x
# LmNybDA4oDagNIYyaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNz
# dXJlZElEQ0EtMS5jcmwwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRw
# Oi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3
# DQEBBQUAA4IBAQCdJX4bM02yJoFcm4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8
# WhY5AJ3jbITkWkD73gYBjDf6m7GdJH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkF
# azWQEKB7l8f2P+fiEUGmvWLZ8Cc9OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+Nt
# FxMG7uQDthSr849Dp3GdId0UyhVdkkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT
# 5VUW8LsKqxzbXEgnZsijiwoc5ZXarsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ir
# a8JRwgpIr7DUbuD0FAo6G+OPPcqvao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YD
# reoACus/J7u6GzANBgkqhkiG9w0BAQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYD
# VQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAw
# WhcNMjExMTEwMDAwMDAwWjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdp
# Q2VydCBBc3N1cmVkIElEIENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
# AoIBAQDogi2Z+crCQpWlgHNAcNKeVlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561
# JBXCmLm0d0ncicQK2q/LXmvtrbBxMevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFI
# gHweog+SDlDJxofrNj/YMMP/pvf7os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C
# 6Ds97bFBo+0/vtuVSMTuHrPyvAwrmdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4G
# rq2Eb0iEm09AufFM8q+Y+/bOQF1c9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrho
# frfVdwonVnwPYqQ/MhRglf0HBKIJAgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMC
# AYYwOwYDVR0lBDQwMgYIKwYBBQUHAwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYB
# BQUHAwQGCCsGAQUFBwMIMIIB0gYDVR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEE
# MIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNw
# cy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1
# AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABj
# AG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBm
# ACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBk
# ACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBl
# AG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABp
# AHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAg
# AGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglghkgB
# hv1sAxUwEgYDVR0TAQH/BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4E
# FgQUFQASKxOYspkH7R7for5XDStnAs0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGL
# p6chnfNtyA8wDQYJKoZIhvcNAQEFBQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFH
# q1Zce9UNC0Gz7+x1H3Q48rJcYaKclcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR
# 67s2hHfMJKXzBBlVqefj56tizfuLLZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPk
# mzmnxPXOHXh2lCVz5Cqrz5x2S+1fwksW5EtwTACJHvzFebxMElf+X+EevAJdqP77
# BzhPDcZdkbkPZ0XN1oPt55INjbFpjE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718
# CqDuLAhVhSK46xgaTfwqIa1JMYNHlXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQe
# MIIEGgIBATBqMFoxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52
# LXNhMTAwLgYDVQQDEydHbG9iYWxTaWduIENvZGVTaWduaW5nIENBIC0gU0hBMjU2
# IC0gRzMCDF8qMMA1ngrijFda+DAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEK
# MAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3
# AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUiUi9N6BhRaRFkZ4D
# AANwmcQzcqswDQYJKoZIhvcNAQEBBQAEggEAQ/eBcSESGKagvhtqYFVDKGgVhovH
# cKpINrjqOtXXLMryRiOtlAtSUUSirNaGsb4A9h581+W2A/OYKVwKorBcgPuJ78n3
# QKYehluol4y1fA0R6p3k+V0Tdd8QBU/1nV7eZboXRqF4u6ny7/qW9JpYBp8C6voA
# OWEvJQQOEh4pelJDeHOtxg9UAt4PXbLorFOEnkiOjX0JwAImTFVow5wldOC+ApKR
# BrB4WcyLBo1Y3wDkzcLQhCUVgFArbwhzxdKODGafPZUIwLbBcEtE/tMV4B6mZb7Z
# ED+KlnVAvM527uILxQzwobf/yigmQVGbBtvMXVNe1BnXRSJxF3ObHOTzvqGCAg8w
# ggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfw
# ZjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMjAwMTAyMTMyODA5WjAjBgkqhkiG9w0BCQQxFgQUh+rYk7145WxN
# +EnCxkAWHNGb8wIwDQYJKoZIhvcNAQEBBQAEggEAIdGX4ZYxlvjHRfj8TJY5Sq4a
# m3rjW18NNpw5TBhFYk9BW2GJEGoMZzA0IeZA0pW9dfAr+icziUhkHJ0ZnHa7q7bd
# zroYUniB46/RRxesi98VjjcBG14qGMkeKtepmGqweWISC7vfv2za9Q1Id7CKxJyc
# 2+jPxJSub7B616KTH/k4fLHgMtGG8/QLMUxN4fUkY8dqchSP7ON1h0iWA8JBqa0f
# iv1neCRINhVy2IpycpmsHtDvUJTFVchjhmqmE0ahhVWMlMKn/gvnGdEBf2BZ3hd8
# vnpP/uz8NsoQt2mazS7iHVAxNrxm/cMw7p6KIFyBoiCwFWUNwZgT9gxuGtpSBQ==
# SIG # End signature block
