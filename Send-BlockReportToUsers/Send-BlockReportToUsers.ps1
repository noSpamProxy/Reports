<#
.SYNOPSIS
  Name: Send-Send-BlockReportToUsers.ps1
  Create report of permanently blocked emails for your employees.

.DESCRIPTION
  This script can be used to generate a report about E-Mails which where permanently blocked.
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
  Default: Auswertung der abgewiesenen E-Mails an Sie
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
  Version:        1.0.1
  Author:         Jan Jaeschke
  Creation Date:  2020-01-02
  Purpose/Change: added validation pattern for FromDate and ToDate parameters
  
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
		[string] $ReportSubject = "Auswertung der abgewiesenen E-Mails an Sie",
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
		<title>Abgewiesene E-Mails an Sie</title>
		<style>
			table, td, th { border: 1px solid black; border-collapse: collapse; padding:10px; text-align:center;}
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
		$messageTracks = Get-NSPMessageTrack @cleanedParams -Status PermanentlyBlocked -WithAddresses -Directions FromExternal -First 100
	}else{
		$messageTracks = Get-NSPMessageTrack @cleanedParams -Status PermanentlyBlocked -WithAddresses -Directions FromExternal -First 100 -Skip $skipMessageTracks
	}

	foreach ($messageTrack in $messageTracks){
		$addresses = $messageTrack.Addresses
		foreach ($addressEntry in $addresses){
			if ($addressEntry.AddressType -eq "Recipient"){
				$messageRecipient = $addressEntry.Address
				if($reportAll -eq $false){
					if ($messageRecipient -notin $uniqueReportRecipientList){
						continue
					}
				}
				<# 
					create tmp list containing the data of hashtable "entries" for the key "messageRecipient"
					if there is no data use the current messagetrack else add the current messagetrack to the data
					save the  tmp list back into the hashtable for the used key 
				#>
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
		$htmlContent = ""
        $_.Name
        foreach ($validationItem in $_.Value) {
            $NSPStartTime = $validationItem.Sent.LocalDateTime
            $addresses2 = $validationItem.Addresses
            $NSPSender = ($addresses2 | Where-Object { $_.AddressType -eq "Sender" } | Select-Object "Address").Address		
            $NSPSubject = $validationItem.Subject
            $htmlContent += "<tr><td width=150px>$NSPStartTime</td><td>$NSPSender</td><td>$NSPSubject</td></tr>`r`n`t`t`t"
        }
        createHTML $htmlContent
        Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $_.Name -Subject $ReportSubject -BodyAsHtml -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der Auswertung der abgewiesenen E-Mails." -Attachments $reportFile
	}
}else{
	Write-Output "Nothing found for report generation."
}

if(Test-Path $reportFile -PathType Leaf){
	Write-Output "Doing some cleanup...."
	Remove-Item $reportFile
}
# SIG # Begin signature block
# MIIYkwYJKoZIhvcNAQcCoIIYhDCCGIACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU3N/wJTor9ObS+Ui23+3GLTJg
# JNGgghPOMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
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
# /xrNfYi4zYch/b/ZtjCCBP4wggPmoAMCAQICEA1CSuC+Ooj/YEAhzhQA8N0wDQYJ
# KoZIhvcNAQELBQAwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IElu
# YzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQg
# U0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQTAeFw0yMTAxMDEwMDAwMDBa
# Fw0zMTAxMDYwMDAwMDBaMEgxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjEgMB4GA1UEAxMXRGlnaUNlcnQgVGltZXN0YW1wIDIwMjEwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDC5mGEZ8WK9Q0IpEXKY2tR1zoRQr0K
# dXVNlLQMULUmEP4dyG+RawyW5xpcSO9E5b+bYc0VkWJauP9nC5xj/TZqgfop+N0r
# cIXeAhjzeG28ffnHbQk9vmp2h+mKvfiEXR52yeTGdnY6U9HR01o2j8aj4S8bOrdh
# 1nPsTm0zinxdRS1LsVDmQTo3VobckyON91Al6GTm3dOPL1e1hyDrDo4s1SPa9E14
# RuMDgzEpSlwMMYpKjIjF9zBa+RSvFV9sQ0kJ/SYjU/aNY+gaq1uxHTDCm2mCtNv8
# VlS8H6GHq756WwogL0sJyZWnjbL61mOLTqVyHO6fegFz+BnW/g1JhL0BAgMBAAGj
# ggG4MIIBtDAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDBBBgNVHSAEOjA4MDYGCWCGSAGG/WwHATApMCcGCCsGAQUF
# BwIBFhtodHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwHwYDVR0jBBgwFoAU9Lbh
# IB3+Ka7S5GGlsqIlssgXNW4wHQYDVR0OBBYEFDZEho6kurBmvrwoLR1ENt3janq8
# MHEGA1UdHwRqMGgwMqAwoC6GLGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEy
# LWFzc3VyZWQtdHMuY3JsMDKgMKAuhixodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# c2hhMi1hc3N1cmVkLXRzLmNybDCBhQYIKwYBBQUHAQEEeTB3MCQGCCsGAQUFBzAB
# hhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTwYIKwYBBQUHMAKGQ2h0dHA6Ly9j
# YWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1cmVkSURUaW1lc3Rh
# bXBpbmdDQS5jcnQwDQYJKoZIhvcNAQELBQADggEBAEgc3LXpmiO85xrnIA6OZ0b9
# QnJRdAojR6OrktIlxHBZvhSg5SeBpU0UFRkHefDRBMOG2Tu9/kQCZk3taaQP9rhw
# z2Lo9VFKeHk2eie38+dSn5On7UOee+e03UEiifuHokYDTvz0/rdkd2NfI1Jpg4L6
# GlPtkMyNoRdzDfTzZTlwS/Oc1np72gy8PTLQG8v1Yfx1CAB2vIEO+MDhXM/EEXLn
# G2RJ2CKadRVC9S0yOIHa9GCiurRS+1zgYSQlT7LfySmoc0NR2r1j1h9bm/cuG08T
# HfdKDXF+l7f0P4TrweOjSaH6zqe/Vs+6WXZhiV9+p7SOZ3j5NpjhyyjaW4emii8w
# ggUxMIIEGaADAgECAhAKoSXW1jIbfkHkBdo2l8IVMA0GCSqGSIb3DQEBCwUAMGUx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9v
# dCBDQTAeFw0xNjAxMDcxMjAwMDBaFw0zMTAxMDcxMjAwMDBaMHIxCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBUaW1lc3Rh
# bXBpbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC90DLuS82P
# f92puoKZxTlUKFe2I0rEDgdFM1EQfdD5fU1ofue2oPSNs4jkl79jIZCYvxO8V9PD
# 4X4I1moUADj3Lh477sym9jJZ/l9lP+Cb6+NGRwYaVX4LJ37AovWg4N4iPw7/fpX7
# 86O6Ij4YrBHk8JkDbTuFfAnT7l3ImgtU46gJcWvgzyIQD3XPcXJOCq3fQDpct1Hh
# oXkUxk0kIzBdvOw8YGqsLwfM/fDqR9mIUF79Zm5WYScpiYRR5oLnRlD9lCosp+R1
# PrqYD4R/nzEU1q3V8mTLex4F0IQZchfxFwbvPc3WTe8GQv2iUypPhR3EHTyvz9qs
# EPXdrKzpVv+TAgMBAAGjggHOMIIByjAdBgNVHQ4EFgQU9LbhIB3+Ka7S5GGlsqIl
# ssgXNW4wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wEgYDVR0TAQH/
# BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgw
# eQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0
# dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5j
# cmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3Vy
# ZWRJRFJvb3RDQS5jcmwwUAYDVR0gBEkwRzA4BgpghkgBhv1sAAIEMCowKAYIKwYB
# BQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCwYJYIZIAYb9bAcB
# MA0GCSqGSIb3DQEBCwUAA4IBAQBxlRLpUYdWac3v3dp8qmN6s3jPBjdAhO9LhL/K
# zwMC/cWnww4gQiyvd/MrHwwhWiq3BTQdaq6Z+CeiZr8JqmDfdqQ6kw/4stHYfBli
# 6F6CJR7Euhx7LCHi1lssFDVDBGiy23UC4HLHmNY8ZOUfSBAYX4k4YU1iRiSHY4yR
# UiyvKYnleB/WCxSlgNcSR3CzddWThZN+tpJn+1Nhiaj1a5bA9FhpDXzIAbG5KHW3
# mWOFIoxhynmUfln8jA/jb7UBJrZspe6HUSHkWGCbugwtK22ixH67xCUrRwIIfEmu
# E7bhfEJCKMYYVs9BNLZmXbZ0e/VWMyIvIjayS6JKldj1po5SMYIELzCCBCsCAQEw
# ajBaMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEwMC4G
# A1UEAxMnR2xvYmFsU2lnbiBDb2RlU2lnbmluZyBDQSAtIFNIQTI1NiAtIEczAgxf
# KjDANZ4K4oxXWvgwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKEC
# gAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwG
# CisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFCkXwjAtFWOGKawRSfwUu44/3Ai1
# MA0GCSqGSIb3DQEBAQUABIIBAAhstUD4pd6ekLL9+T77LT5WEKQPiFb7gzxrzxM5
# VUI7PPc8wWmrFUi+r6NkxOdIsbhdN2wpAK/u8AsIHVAeBizi4ZdNenBFJKeLSikX
# iPsdFD0+lUBXVmxbKExy+l3XKl9kGXkgkE/M6H7YhTIM/gB/apbjLBJEgEVjNKaI
# pLPTCDbae6vHv3830FiKQsTVg6279j798v5JqTQm6ay9Z7d1TrqXXmDyiZpX7BQP
# b22niaM0bLyWQoqESOjTeoz4XURpx/AIr/iGuzspQNAGSeKyHZ0JkdQnsFnQIw9n
# diBlYMqTePy6LyvLyVkH+DVP/js88iT1SNPfnkockBCTqm6hggIgMIICHAYJKoZI
# hvcNAQkGMYICDTCCAgkCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERp
# Z2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMo
# RGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQQIQDUJK4L46
# iP9gQCHOFADw3TAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEH
# ATAcBgkqhkiG9w0BCQUxDxcNMjEwMTA1MTM0NTEzWjAjBgkqhkiG9w0BCQQxFgQU
# OslGjF9pR14FCpHjaDlZQn2imU0wDQYJKoZIhvcNAQEBBQAEggEAt4RGV8a/zJz1
# ZlH0XpzLKZhwmAV8UPhELIld4MDeRisgYmn5WX/eGbbvKXgUhbraLhYNgA88HkUg
# ZMVx8NGEzrmR+OYHFNu2N8O1ndjq4YOV4gaAjshrHtsMbbwTfDY8dpqYiXhlolqj
# LSXiYJCvGO79FasEZDs9E5FhlDv8ixmP3oja67PZOABoU+IIadkx6QnyL46Q6ZgD
# /UZz80Aaa5w6TjtxUIfTlZ5QxqggRhes88sFFBGCHmk+LR8niARTmaijjT/pzHsZ
# XUpjBdx8qusNzcRm5xVZBeWwzxviNXE9K3te4QpUj+yGp4BL4IwOHVGiSwaXkO2S
# gWFvUmNmOg==
# SIG # End signature block
