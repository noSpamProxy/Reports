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

 .PARAMETER CheckuserExistance
  The switch allows to check each report recipient against the known NoSpamProxy users.
  Only usable if no recipient list is provided.
  Can have a huge performance impact.

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

.PARAMETER TenantPrimaryDomain
  Used to login into the desired NoSpamProxy tenant to run this script on.
  Only required if NoSpamProxy v14 is used in provider mode.
	
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
  Version:        1.1.1
  Author:         Jan Jaeschke
  Creation Date:  2024-09-16
  Purpose/Change: added seperate temp folder for report file for MSP and added recipient address to email body
  
.LINK
  https://www.nospamproxy.de
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
	# enable user existance check
	[parameter(ParameterSetName = "dateSpanSet")][parameter(ParameterSetName = "numberOfDaysSet")][parameter(ParameterSetName = "numberOfHoursSet")][parameter(ParameterSetName = "noTimeSet")][parameter(ParameterSetName = "reportIntervalSet")]
		[switch] $CheckUserExistence,
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
		[string] $AdUsername,
	# only needed for v14 with enabled provider mode
	[Parameter(Mandatory=$false)][string] $TenantPrimaryDomain	
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
	$FromDate = Get-Date -Date ((Get-Date).AddHours(-$NumberOfHoursToReport)) -UFormat "%Y-%m-%d %H:%M:%S"
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
if ($TenantPrimaryDomain) {
	New-Item -ItemType Directory -Path $ENV:Temp + "\" + "$TenantPrimaryDomain"
	$reportFile = $Env:TEMP + "\" + "$TenantPrimaryDomain" + "$ReportFileName" + ".html"
} else {
	$reportFile = $Env:TEMP + "\" + "$ReportFileName" + ".html"
}

$entries = @{}


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
				} elseif ($CheckUserExistence) {
					if (!(Get-Nspuser -Filter "$messageRecipient")) {
						Write-Verbose "$messageRecipient is not a known nsp user."
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
		$htmlContent = "Recipient Address: $($_.Name)"
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
	if ($TenantPrimaryDomain) {
		Remove-Item $ENV:Temp + "$TenantPrimaryDomain"
	}
}