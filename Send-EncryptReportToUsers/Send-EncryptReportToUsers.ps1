<#
.SYNOPSIS
  Name: Send-EncryptionReportToUsers.ps1
  Create report of outbound emails for your employees.

.DESCRIPTION
  This script can be used to generate a report about outbound E-Mails and shows if they where encrypted or not.
  It is possible to filter the results for a specific time duration and sends the report only to specific or all affected users.
  This script only uses the NoSpamProxy Powershell Cmdlets to generate the report file.

.PARAMETER AdBaseDN
  Define the BaseDN for searching a user group in the defined AD.

.PARAMETER AdPort
  Default: 389
  Define a custom port to access the AD.

.PARAMETER AdReportGroup
  Define the AD user group to search for.
  The users in this group will receive a report.

.PARAMETER AdServer
  Define the hostname, FQDN or IP address of the desired AD.

.PARAMETER AdUsername
  Define an optional username to authenticate against the AD.
  A password have to be set before using <SetAdPassword>.

.PARAMETER FromDate
  Mandatory if you like to use a timespan.
  Specifies the start date for the E-Mail filter.
  Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"
  E.g.:
  	"2019-06-05 08:00" or "2019-06-05 20:00:00"

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
  Default: reject-analysis
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
  Default: Auswertung der Ihrer versendeten E-Mails
  Sets the report E-Mail subject.

.PARAMETER SetAdPassword
  Set a password to use for authentication against the AD.
  The password will be saved as encrypted "NspAdReportingPass.bin" file under %APPDATA%.

.PARAMETER SmtpHost
  Default: 127.0.0.1
  Specifies the SMTP host which should be used to send the report E-Mail.
  It is possible to use a hostname, FQDN or IP address.
	
.PARAMETER ToDate
  Optional if you like to use a timespan.
  Specifies the end date for the E-Mail filter.
  Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"
  E.g.:
	  "2019-06-05 08:00" or "2019-06-05 20:00:00"
	  
.OUTPUTS
  Report is stored under %TEMP%\reject-analysis.html unless a custom <ReportFileName> parameter is given.
  Will be deleted after the email to the recipient was sent.

.NOTES
  Version:        1.0.0
  Author:         Jan Jaeschke
  Creation Date:  2020-01-02
  Purpose/Change: inital script creation
  
.LINK
  https://https://www.nospamproxy.de
  https://github.com/noSpamProxy

.EXAMPLE
  .\Send-Send-BlockReportToUsers.ps1 -NoTime -ReportFileName "Example-Report" -ReportRecipient alice@example.com -ReportSender "NoSpamProxy Report Sender <nospamproxy@example.com>" -ReportSubject "Example Report" -SmtpHost mail.example.com
  
.EXAMPLE
  .\Send-Send-BlockReportToUsers.ps1 -FromDate: "2019-06-05 08:00:00" -ToDate: "2019-06-05 20:00:00" 
  It is mandatory to specify <FromDate>. Instead <ToDate> is optional.
  These parameters can be combined with all other parameters except <NumberOfDaysToReport>, <NumberOfHoursToRepor>, <ReportIntervall> and <NoTime>.

.EXAMPLE 
  .\Send-Send-BlockReportToUsers.ps1 -NumberOfDaysToReport 7 
  You can combine <NumberOfDaysToReport> with all other parameters except <FromDate>, <ToDate>, <NumberOfHoursToRepor>, <ReportIntervall> and <NoTime>.
  
.EXAMPLE 
  .\Send-Send-BlockReportToUsers.ps1 -NumberOfHoursToReport 12
  You can combine <NumberOfHoursToReport> with all other parameters except <FromDate>, <ToDate>, <NumberOfDaysToReport>, <ReportIntervall> and <NoTime>.
	
.EXAMPLE
	.\Send-BlockReportToUsers.ps1 -ReportInterval weekly
	You can combine <ReportInterval> with all other parameters except <FromDate>, <ToDate>, <NumberOfDaysToReport>, <NumberOfHoursToReport>, <ReportIntervall> and <NoTime>.

.EXAMPLE
  .\Send-BlockReportToUsers.ps1 -NoTime -NspRule "All other inbound mails"
  
.EXAMPLE
  .\Send-BlockReportToUsers.ps1 -NoTime -SmtpHost mail.example.com -ReportRecipientCSV "C:\Users\example\Documents\email-report.csv"
  The CSV have to contain the header "Email" else the mail addresses cannot be read from the file.
  It is possible to combine <ReportRecipientCSV> with <ReportRecipient> and a AD group.
  E.g: email-report.csv
  User,Email
  user1,user1@example.com
  user2,user2@example.com

.EXAMPLE
  .\Send-BlockReportToUsers.ps1 -NoTime -AdServer ad.example.com -AdBaseDN "DC=example,DC=com" -AdReportGroup "MyReportGroup"
  Connect to AD as anonymous.

.EXAMPLE 
  .\Send-BlockReportToUsers.ps1 -NoTime -AdServer ad.example.com -AdBaseDN "DC=example,DC=com" -AdReportGroup "MyReportGroup" -AdUsername Administrator
  Connect to AD as Administrator, password needs to be set using .\Send-BlockReportToUsers.ps1 -SetAdPassword

.EXAMPLE 
  .\Send-BlockReportToUsers.ps1 -SetAdPassword
  Will wait for a user input. The input is shown in plain text!
#>
Param(
	# set start date for filtering
	[Parameter(Mandatory=$true, ParameterSetName="dateSpanSet")]
		[ValidatePattern("^([0-9]{4})-?(1[0-2]|0[1-9])-?(3[01]|0[1-9]|[12][0-9])\s(2[0-3]|[01][0-9]):?([0-5][0-9]):?([0-5][0-9])?$")]
		[string] $FromDate,
	# set end date for filtering
	[Parameter(Mandatory=$false, ParameterSetName="dateSpanSet")]
		[ValidatePattern("^([0-9]{4})-?(1[0-2]|0[1-9])-?(3[01]|0[1-9]|[12][0-9])\s(2[0-3]|[01][0-9]):?([0-5][0-9]):?([0-5][0-9])?$")]
		[string] $ToDate,
	# if set not time duration have to be set the E-Mail of the last few hours will be filtered
	[Parameter(Mandatory=$true, ParameterSetName="noTimeSet")]
		[switch] $NoTime, 
	# set number of days for filtering
	[Parameter(Mandatory=$true, ParameterSetName="numberOfDaysSet")]	
		[ValidatePattern("[0-9]+")]
		[string] $NumberOfDaysToReport,
	# set number of hours for filtering
	[Parameter(Mandatory=$true, ParameterSetName="numberOfHoursSet")]
		[int] $NumberOfHoursToReport,
	# set reporting intervall
	[Parameter(Mandatory=$true, ParameterSetName="reportIntervalSet")]
		[ValidateSet('daily','weekly','monthly')]
		[string] $ReportInterval,
	# run script to save AD user password in an encrypted file
	[Parameter(Mandatory=$true, ParameterSetName="setAdPassword")]
		[switch] $SetAdPassword,
	# set report sender address for outbound email
	[parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[ValidatePattern("^([a-zA-Z0-9\s.!£#$%&'^_`{}~-]+)?<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*>?$")]
		[string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",	
	# set outbound email subject
	[parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[string] $ReportSubject = "Auswertung Ihrer versendeten E-Mails",
	# change report filename
	[Parameter(Mandatory=$false)]
		[string] $ReportFileName = "reject-analysis",
	# set smtp host for relaying outpund email
	[parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[ValidatePattern("^(((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))|((([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])))$")]
		[string] $SmtpHost = "127.0.0.1",
	# set report recipient only valid E-Mail addresses are allowed
	[Parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[ValidatePattern("^<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+?<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*>?$")]
		[string[]] $ReportRecipient,
	# set path to csv file containing report recipient E-Mail addresses
	[Parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[string] $ReportRecipientCSV,
	# set AD host
	[parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[ValidatePattern("^(((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))|((([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])))$")]
		[string] $AdServer,	
	# set port to access AD
	[parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[ValidateRange(0,65535)]
		[int] $AdPort = 389,
	# set base DN for filtering
	[parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[string] $AdBaseDN,
	# set AD security group containing the desired user objects
	[parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[string] $AdReportGroup,
	# set  AD username for authorization
	[parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[string] $AdUsername
)

#-------------------Functions----------------------
# save AD password as encrypted file
function Set-adPass{
	$adPass = Read-Host -Promp 'Input your AD User password'
	$passFileLocation = $(Join-Path $env:APPDATA 'NspAdReportingPass.bin')
    $inBytes = [System.Text.Encoding]::Unicode.GetBytes($adPass)
    $protected = [System.Security.Cryptography.ProtectedData]::Protect($inBytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
	[System.IO.File]::WriteAllBytes($passFileLocation, $protected)
}
# read encrypted AD password from file if existing else user is promted to enter the password
function Get-adPass {
	$passFileLocation = $(Join-Path $env:APPDATA 'NspAdReportingPass.bin')
    if (Test-Path $passFileLocation) {
        $protected = [System.IO.File]::ReadAllBytes($passFileLocation)
        $rawKey = [System.Security.Cryptography.ProtectedData]::Unprotect($protected, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::Unicode.GetString($rawKey)
    } else {
		Write-Output "No Password file found! Please run 'Send-ReportToUsersWithBlockedEmails.ps1 -SetPassword' for saving your password encrypted."
		$adPass = Read-Host -Promp 'Input your AD User password'
		return $adPass
    }
}
# generate HMTL report
function createHTML($htmlContent) {
	$htmlOut =
"<html>
	<head>
		<title>Verschlüsselte E-Mails von Ihnen</title>
		<style>
			table, td, th { border: 1px solid black; border-collapse: collapse; padding:10px; text-align:center;}
			#headerzeile         {background-color: #DDDDDD;}
		</style>
	</head>
	<body style=font-family:arial>
		<h1>Verschlüsselte E-Mails von Ihnen</h1>
		<br>
		<table>
			 <tr id=headerzeile>
			 <td><h3>Uhrzeit</h3></td><td><h3>Empfänger</h3></td><td><h3>Betreff</h3></td><td><h3>Encrypted</h3></td>
			 </tr>
			 $htmlContent				
		</table>
	</body>
</html>"
	$htmlOut | Out-File "$reportFile"

}

#-------------------Variables----------------------
# check SmtpHost is set
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
}elseif ($NumberOfHoursToReport){
	$FromDate = (Get-Date).AddHours(-$NumberOfHoursToReport)
}

# create hashtable which will preserve the order of the added items and mapps userParams into needed parameter format
$userParams = [ordered]@{ 
	From = $FromDate
	To = $ToDate
	Age = $NumberOfDaysToReport
	Rule = $NspRule
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
$reportAll = $false

#--------------------Main-----------------------
# set AD User password and exit program
if($SetAdPassword){
	# Imports Security library for encryption
	Add-Type -AssemblyName System.Security
	Set-adPass
	EXIT
}

if($AdServer){
	# Imports Security library for encryption
	Add-Type -AssemblyName System.Security
	# create AD connection
	# create ADSISearcher object
		$ds=[AdsiSearcher]""
	# define the needed AD object properties
	$ds.PropertiesToLoad.AddRange(@('mail'))
	# define AD search filter
	$ds.filter="(&((memberOf=CN=$AdReportGroup,$AdBaseDN)(ObjectCategory=user)))"
	# define AD paging
	$ds.pagesize=100
	# check if username and read the password from encrypted file else use anonymous connection
	if($AdUsername){
		$password = Get-adPass
		$ds.searchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://${AdServer}:$AdPort","$AdUsername","$password")
	}else{
		$ds.searchRoot = [Adsi]"LDAP://${AdServer}:$AdPort"
	}
	# get all desired users from AD
	$UserListe = $ds.findall()
	# save users mail addresses
	$userMailList = $UserListe.Properties.mail
}

if($ReportRecipientCSV){
	$csv = Import-Csv $ReportRecipientCSV
	$csvMailRecipient = $csv.Email
}
$uniqueReportRecipientList = (($ReportRecipient + $csvMailRecipient + $userMailList) | Get-Unique)
if($null -eq $uniqueReportRecipientList){
	Write-Output "No report recipient list generated, every affected user will receive a report."
	$reportAll = $true
}

# condition to run Main part, if false program will end
$getMessageTracks = $true
$skipMessageTracks = 0
$reportFile = $Env:TEMP + "\" + "$ReportFileName" + ".html"
$entries = @{}


#--------------------Main-----------------------
while($getMessageTracks -eq $true){	
	if($skipMessageTracks -eq 0){
		$messageTracks = Get-NSPMessageTrack @cleanedParams -Status Success -WithAddresses -WithOperations -Directions FromLocal -First 100
	}else{
		$messageTracks = Get-NSPMessageTrack @cleanedParams -Status Success -WithAddresses -WithOperations -Directions FromLocal -First 100 -Skip $skipMessageTracks
	}
	foreach ($messageTrack in $messageTracks){
		$addresses = $messageTrack.Addresses
		foreach ($addressEntry in $addresses){
			if ($addressEntry.AddressType -eq "Sender"){
				$messageSender = $addressEntry.Address
				if($reportAll -eq $false){
					if ($messageSender -notin $uniqueReportRecipientList){
						continue
					}
				}
				<# 
					create tmp list containing the data of hashtable "entries" for the key "messageRecipient"
					if there is no data use the current messagetrack else add the current messagetrack to the data
					save the  tmp list back into the hashtable for the used key 
				#>
				$list = $entries[$messageSender]
				if (!$list) {
					$list = @($messagetrack)
				}
				else
				{
					$list += $messageTrack
				}
				$entries[$messageSender] = $list
				}
		}
	}
	# exit condition
	if($messageTracks){
		$skipMessageTracks = $skipMessageTracks+100
		Write-Verbose $skipMessageTracks
	}else{
		$getMessageTracks = $false
		break
	}
}
if($entries.Count -ne 0){
    Write-Output "Generating and sending reports for the following e-mail addresses:"
    $entries.GetEnumerator() | ForEach-Object {
		$_.Name
		$htmlContent = ""
        foreach ($validationItem in $_.Value) {
			$isEncrypted = "no"
			$messageOperations = $validationItem.Operations
			foreach($operation in $messageOperations){
				if($operation.Operation.Type -eq "SMimeEncryption"){
					$isEncrypted = "yes"
					continue
				}elseif($operation.Operation.Type -eq "PdfEncryption"){
					$isEncrypted = "yes"
					continue
				}
			}
			$NSPStartTime = $validationItem.Sent.LocalDateTime
			$addresses2 = $validationItem.Addresses
			$NSPRecipient = ($addresses2 | Where-Object { $_.AddressType -eq "Recipient" } | Select-Object "Address").Address		
			$NSPSubject = $validationItem.Subject
			$htmlContent += "<tr><td width=150px>$NSPStartTime</td><td>$NSPRecipient</td><td>$NSPSubject</td><td>$isEncrypted</td></tr>`r`n`t`t`t"
		}
		createHTML $htmlContent
		Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $_.Name -Subject $ReportSubject -BodyAsHtml -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der Auswertung der von Ihnen versendeten E-Mails." -Attachments $reportFile
	}
}else{
	Write-Output "Nothing found for report generation."
}

if(Test-Path $reportFile -PathType Leaf){
	Write-Output "Doing some cleanup...."
	Remove-Item $reportFile
}
# SIG # Begin signature block
# MIIbigYJKoZIhvcNAQcCoIIbezCCG3cCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUVHR1cjVjasX5TtpQylz+Q4Lq
# v9WgghbWMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
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
# AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUw6NvlTvsiKA7v49v
# qqOdozntBqkwDQYJKoZIhvcNAQEBBQAEggEAf1XZCqtNx2bxb1aDt63raZMYjXKf
# 4cQn7Fac1izawDhWpA7tkDTs2sDUe78f/NCpiW8aEd8RTF2OHh8bayuwvFAgcspY
# 8x/HOh6H1K/LMEu0vXW1hwZ+ijZ2r/5/c1zkY2K0GK4nPCrzu4NmE3hskoZgURHc
# nAOVVgewZJ0IV79IblJPgGdiTgl3JilkfSjpCd/0TX7siUVRf5Iabh8bt4TpfD4c
# 7EQSho5v6dTpCAYKDveRJZ8qBYsxgpgOXMJ6cxhcXaTH2WpO5adMkzyTtd3RdrRJ
# rU12P2vI9WEzs0haG4DcJCweZCMtb0hxnUZ7bzedIUoaAUVNGuAJN42upaGCAg8w
# ggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfw
# ZjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMjAwMTAyMTQ1MzA4WjAjBgkqhkiG9w0BCQQxFgQUSBK4qCU6qvfO
# zaPLZSplVRMry3AwDQYJKoZIhvcNAQEBBQAEggEAeW1CJNc8AceuaAUDOh4Nc/aS
# PWI8VzVsWvWdxxD+9RmDaU2TuNJrC61KobK156IOt9IQoa+DungdMbe9aAH5Caw5
# QOaG/1SwXBUqJzadrd/gaIEF2QYL8HJjDKJYNhoF1qnCCtXmURz+pGALbUsHuKqV
# DAF7NomuMDYxV0HtS9K+AsfeapVfnsIi3fDAP071haFyuI0d5I3Wmy18dEYzWUbZ
# GfFXUUdEZMgU8+IXEG/EXF7PioXofEVhUyc9txWw/YhyMLzc7OidIsVBkyTOD66e
# q7fh5p3XX20NncbOYDskCJQudLVypCWRq4AVPpLaY5k27UuEgetMdfJZwmcgWQ==
# SIG # End signature block
