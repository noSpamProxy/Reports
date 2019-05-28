<#
.SYNOPSIS
  Name: Get-NspLicenseReport.ps1
  Create NSP license report from the last 90 days.

.DESCRIPTION
  This script can be used to generate a license report either user-based, domain-based or both.
  It is possible to send the report via E-Mail to one or multiple recipients.
	This script uses the NoSpamProxy Powershell Cmdlets and an SQL query to generate the report files.
	The report will be generated always for the past 90 days.

.PARAMETER NoMail
	Does not generate an E-Mail and saves the genereated reports to the current execution location.
	Ideal for testing or manual script usage.

.PARAMETER ReportFileName
  Default: License_Report
	Define a part of the complete file name.
	E.g.: 
	user-based:
	C:\Users\example\Documents\2019-05-27_License_Report_per_user.txt
	domain-based:
	C:\Users\example\Documents\2019-05-27_License_Report_example.com.txt
	
.PARAMETER ReportRecipient
  Specifies the E-Mail recipient. It is possible to pass a comma seperated list to address multiple recipients. 
  E.g.: alice@example.com,bob@example.com

.PARAMETER ReportRecipientCSV
  Set a filepath to an CSV file containing a list of report E-Mail recipient. Be aware about the needed CSV format, please watch the provided example.

.PARAMETER ReportSender
  Default: NoSpamProxy Report Sender <nospamproxy@example.com>
  Sets the report E-Mail sender address.
  
.PARAMETER ReportSubject
  Default: NSP License Report
	Sets the report E-Mail subject.
	
.Parameter ReportType
	Default: user-based
	Sets the type of generated report.
	Possible values are: user-based, domain-based, both

.PARAMETER SmtpHost
  Specifies the SMTP host which should be used to send the report E-Mail.
	It is possible to use a FQDN or IP address.
	
.PARAMETER SqlCredential
	Sets custom credentials for database access.
	By default the authentication is done using current users credentials from memory.

.PARAMETER SqlDatabase
	Default: NoSpamProxyAddressSynchronization
	Sets a custom SQl database name which should be accessed. The required database is the one from the intranet-role.

.PARAMETER SqlInstance
	Default: NoSpamProxy
	Sets a custom SQL instance name which should be accessed. The required instance must contain the intranet-role database.

.PARAMETER SqlServer
	Default: (local)
	Sets a custom SQL server which must contains the instance and the database of the intranet-role.

.OUTPUTS
	Report is temporary stored under %TEMP% if the report is send via by E-Mail.
	If the parameter <NoMail> is used the files will be saved at the current location of the executing user.

.NOTES
  Version:        1.0.0
  Author:         Jan Jaeschke
  Creation Date:  2019-05-27
  Purpose/Change: inital creation of the script
  
.LINK
  https://https://www.nospamproxy.de
  https://github.com/noSpamProxy

.EXAMPLE
  .\Get-NspLicenseReport.ps1 -ReportRecipient alice@example.com -ReportSender "NoSpamProxy Report Sender <nospamproxy@example.com>" -ReportSubject "Example Report" -SmtpHost mail.example.com
  
.EXAMPLE
  .\Get-NspLicenseReport.ps1 -ReportRecipient alice@example.com -ReportSender "NoSpamProxy Report Sender <nospamproxy@example.com>" -ReportSubject "Example Report" -SmtpHost mail.example.com -ReportRecipientCSV "C:\Users\example\Documents\email-recipient.csv"
  The CSV have to contain the header "Email" else the mail addresses cannot be read from the file. 
  E.g: email-recipient.csv
  User,Email
  user1,user1@example.com
	user2,user2@example.com
	The "User" header is not necessary.

.EXAMPLE
	.\Get-NspLicenseReport.ps1 -NoMail -ReportType both
	Generates a user-based and a domain-based report which are saved at the current location of execution, here: ".\"

.EXAMPLE
	.\Get-NspLicenseReport.ps1 -NoMail -SqlServer sql.example.com -SqlInstance NSPIntranetRole -SqlDatabase NSPIntranet
	This generates a user-based report. Therefore the script connects to the SQL Server "sql.example.com" and accesses the SQL instance "NSPIntranetRole" which contains the "NSPIntranet" database.

.EXAMPLE 
	.\Get-NspLicenseReport.ps1 -NoMail -SqlInstance ""
	Use the above instance name "" if you try to access the default SQL instance.
	If there is aconnection problem and the NSP configuration shows an empty instance for the intranet-role under "Configuration -> NoSpamProxy components" than this instance example should work.
#>
param (
# userParams are used for filtering
	# generate the report but do not send an E-Mail
	[Parameter(Mandatory=$false)][switch]$NoMail, 
	# change report filename
	[Parameter(Mandatory=$false)][string] $ReportFileName = "License_Report" ,
	# set report recipient only valid E-Mail addresses are allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+?<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*>?$")][string[]] $ReportRecipient,
	# set path to csv file containing report recipient E-Mail addresses
	[Parameter(Mandatory=$false)][string] $ReportRecipientCSV,
	# set report sender address only a valid E-Mail addresse is allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*>?$")][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	# change report E-Mail subject
	[Parameter(Mandatory=$false)][string] $ReportSubject = "NSP License Report",
	[Parameter(Mandatory=$false)][ValidateSet('user-based','domain-based','both')][string] $ReportType = "user-based",
	# set used SMTP host for sending report E-Mail only a valid  IP address or FQDN is allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^(((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))|(((?!-)[a-zA-Z0-9-]{0,62}[a-zA-Z0-9]\.)+[a-zA-Z]{2,63}))$")][string] $SmtpHost,
	# sql credentials
	[Parameter(Mandatory=$false)][pscredential] $SqlCredential,
	# database name
	[Parameter(Mandatory=$false)][string] $SqlDatabase = "NoSpamProxyAddressSynchronization",
	# sql server instance
	[Parameter(Mandatory=$false)][string] $SqlInstance = "NoSpamProxy",
	# sql server
	[Parameter(Mandatory=$false)][string] $SqlServer = "(local)"
)

#-------------------Functions----------------------
# send report E-Mail 
function sendMail($ReportRecipient, $ReportRecipientCSV, $reportAttachment){ 
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
		Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $mailRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie einen automatisch Lizenz-Bericht vom NoSpamProxy" -Attachments $reportAttachment
	}
}

# create database connection
function New-DatabaseConnection() {
	$connectionString = "Server=$SqlServer\$SqlInstance;Database=$SqlDatabase;"
	if ($Credential) {
		$networkCredential = $SqlCredential.GetNetworkCredential()
		$connectionString += "uid=" + $networkCredential.UserName + ";pwd=" + $networkCredential.Password + ";"
	}
	else {
		$connectionString +="Integrated Security=True";
	}
	$connection = New-Object System.Data.SqlClient.SqlConnection
	$connection.ConnectionString = $connectionString
	
	$connection.Open()

	return $connection;
}

# run sql query
function Invoke-SqlQuery([string] $queryName, [bool] $isInlineQuery = $false, [bool] $isSingleResult) {
	try {
		$connection = New-DatabaseConnection
		$command = $connection.CreateCommand()
		if ($isInlineQuery) {
			$command.CommandText = $queryName;
		}
		else {
			$command.CommandText = (Get-Content "$PSScriptRoot\$queryName.sql")
		}
		if ($isSingleResult) {
			return $command.ExecuteScalar();
		}
		else {
			$result = $command.ExecuteReader()
			$table = new-object "System.Data.DataTable"
			$table.Load($result)
			return $table
		}
	}
	finally {
		$connection.Close();
	}
}

# generates user-based report
function userBased($licUsage){
	$userReport = "$reportFile" + "_per_user.txt"
	# get license count per feature
	$protectionLicCount = ($licUsage | Where-Object{$_.Protection -eq 1}).Count
	$encryptionLicCount = ($licUsage | Where-Object{$_.Encryption -eq 1}).Count
	$largeFilesLicCount = ($licUsage | Where-Object{$_.LargeFiles -eq 1}).Count
	$disclaimerLicCount = ($licUsage | Where-Object{$_.Disclaimer -eq 1}).Count
	$sandBoxLicCount = ($licUsage | Where-Object{$_.FilesUploadedToSandbox -eq 1}).Count
	# get users per feature
	$protectionUsers = ($licUsage | Where-Object{$_.Protection -eq 1}).DisplayName
	$encryptionUsers = ($licUsage | Where-Object{$_.Encryption -eq 1}).DisplayName
	$largeFilesUsers = ($licUsage | Where-Object{$_.LargeFiles -eq 1}).DisplayName
	$disclaimerUsers = ($licUsage | Where-Object{$_.Disclaimer -eq 1}).DisplayName
	$sandBoxUsers = ($licUsage | Where-Object{$_.FilesUploadedToSandbox -eq 1}).DisplayName

	# generate formated output
	$stream = [System.IO.StreamWriter] $userReport
	$stream.WriteLine("Protection: $protectionLicCount User")
	$stream.WriteLine("-----------------------------")
	$protectionUsers | ForEach-Object{$stream.WriteLine($_)}
	$stream.WriteLine("`r`n`r`nEncryption: $encryptionLicCount User")
	$stream.WriteLine("-----------------------------")
	$encryptionUsers | ForEach-Object{$stream.WriteLine($_)}
	$stream.WriteLine("`r`n`r`nLargeFiles: $largeFilesLicCount User")
	$stream.WriteLine("-----------------------------")
	$largeFilesUsers | ForEach-Object{$stream.WriteLine($_)}
	$stream.WriteLine("`r`n`r`nDisclaimer: $disclaimerLicCount User")
	$stream.WriteLine("-----------------------------")
	$disclaimerUsers | ForEach-Object{$stream.WriteLine($_)}
	$stream.WriteLine("`r`n`r`nSandBox: $sandBoxLicCount User")
	$stream.WriteLine("-----------------------------")
	$sandBoxUsers | ForEach-Object{$stream.WriteLine($_)}
	$stream.Close()

	return $userReport
}

# generates domain-based report
function domainBased($licUsage){
	$ownDomains = ($licUsage.Domain | Get-Unique)
	foreach($domain in $ownDomains){
		$domainReport = "$reportFile" + "_" + "$domain.txt"
		# get license count per feature
		$protectionLicCount = ($licUsage | Where-Object {$_.Protection -eq 1 -and $_.Domain -eq $domain}).Count
		$encryptionLicCount = ($licUsage | Where-Object {$_.Encryption -eq 1 -and $_.Domain -eq $domain}).Count
		$largeFilesLicCount = ($licUsage | Where-Object {$_.LargeFiles -eq 1 -and $_.Domain -eq $domain}).Count
		$disclaimerLicCount = ($licUsage | Where-Object {$_.Disclaimer -eq 1 -and $_.Domain -eq $domain}).Count
		$sandBoxLicCount = ($licUsage | Where-Object {$_.FilesUploadedToSandbox -eq 1 -and $_.Domain -eq $domain}).Count
		# get users per feature
		$protectionUsers = ($licUsage | Where-Object{$_.Protection -eq 1 -and $_.Domain -eq $domain}).DisplayName
		$encryptionUsers = ($licUsage | Where-Object{$_.Encryption -eq 1 -and $_.Domain -eq $domain}).DisplayName
		$largeFilesUsers = ($licUsage | Where-Object{$_.LargeFiles -eq 1 -and $_.Domain -eq $domain}).DisplayName
		$disclaimerUsers = ($licUsage | Where-Object{$_.Disclaimer -eq 1 -and $_.Domain -eq $domain}).DisplayName
		$sandBoxUsers = ($licUsage | Where-Object{$_.FilesUploadedToSandbox -eq 1 -and $_.Domain -eq $domain}).DisplayName

		# generate formated output
		$stream = [System.IO.StreamWriter] "$domainReport"
		$stream.WriteLine("Protection: $protectionLicCount User")
		$stream.WriteLine("-----------------------------")
		$protectionUsers | ForEach-Object{$stream.WriteLine($_)}
		$stream.WriteLine("`r`n`r`nEncryption: $encryptionLicCount User")
		$stream.WriteLine("-----------------------------")
		$encryptionUsers | ForEach-Object{$stream.WriteLine($_)}
		$stream.WriteLine("`r`n`r`nLargeFiles: $largeFilesLicCount User")
		$stream.WriteLine("-----------------------------")
		$largeFilesUsers | ForEach-Object{$stream.WriteLine($_)}
		$stream.WriteLine("`r`n`r`nDisclaimer: $disclaimerLicCount User")
		$stream.WriteLine("-----------------------------")
		$disclaimerUsers | ForEach-Object{$stream.WriteLine($_)}
		$stream.WriteLine("`r`n`r`nSandBox: $sandBoxLicCount User")
		$stream.WriteLine("-----------------------------")
		$sandBoxUsers | ForEach-Object{$stream.WriteLine($_)}
		$stream.Close()
	}
	return $domainReport
}
#-------------------Variables----------------------
# get the current date for report file name
$reportFileDate = Get-Date -UFormat "%Y-%m-%d"
# define file path of the report file
if ($NoMail){
	$reportFilePath = (Get-Location).Path
} else{
	$reportFilePath = $ENV:TEMP
}
# build the complete default report file path
$reportFile =  "$reportFilePath" + "\" + "$reportFileDate" + "_" + "$ReportFileName"

#--------------------Main-----------------------
$licUsage = Invoke-SqlQuery "LicenseUsage"
if($ReportType){
	switch($ReportType){
		'user-based'{
			$userReport = userBased $licUsage
			# send mail if <NoMail> switch is not used and delete temp report file
			if (!$NoMail){
				sendMail $ReportRecipient $ReportRecipientCSV $userReport
				Remove-Item $userReport
			}
		}
		'domain-based'{
			$domainReport = domainBased $licUsage
			if (!$NoMail){
				sendMail $ReportRecipient $ReportRecipientCSV $domainReport
				Remove-Item $domainReport
			}
		}
		'both'{
			$userReport = userBased $licUsage
			$domainReport = domainBased $licUsage
			if (!$NoMail){
				$bothReports = "$userReport", "$domainReport"
				sendMail $ReportRecipient $ReportRecipientCSV $bothReports
				Remove-Item $userReport
				Remove-Item $domainReport
			}
		}
	}
}
Write-Host "Skript durchgelaufen"
# SIG # Begin signature block
# MIIbnQYJKoZIhvcNAQcCoIIbjjCCG4oCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUM+gsQicvMtRHz9Tuc7z/vAJ1
# ybOgghbpMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
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
# fIxA9gccsK8EBtz4bEaIcpzrTp3DsLlUo7lOl8oUMIIFDjCCA/agAwIBAgIMUfr8
# J+jCyr4Ay7YNMA0GCSqGSIb3DQEBCwUAMFoxCzAJBgNVBAYTAkJFMRkwFwYDVQQK
# ExBHbG9iYWxTaWduIG52LXNhMTAwLgYDVQQDEydHbG9iYWxTaWduIENvZGVTaWdu
# aW5nIENBIC0gU0hBMjU2IC0gRzMwHhcNMTYwNzI4MTA1NjE3WhcNMTkwNzI5MTA1
# NjE3WjCBhzELMAkGA1UEBhMCREUxDDAKBgNVBAgTA05SVzESMBAGA1UEBxMJUGFk
# ZXJib3JuMRkwFwYDVQQKExBOZXQgYXQgV29yayBHbWJIMRkwFwYDVQQDExBOZXQg
# YXQgV29yayBHbWJIMSAwHgYJKoZIhvcNAQkBFhFpbmZvQG5ldGF0d29yay5kZTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJWtx+QDzgovn6AmkJ8UCTNr
# xtFJbRCHKNkfev6k35mMkNlibsVnFxooABDKSvaB21nXojMz63g+KLUEN5S4JiX3
# FKq5h2XahwWHvar/r2HMK2uJZ76360ePhuSZTnkifsxvwNxByQ9ot2S1O40AyVU5
# xfEUsBh7vVADMbjqBVlXuNAfsfpfvgjoR0CsOfgKk0CEDZ1wP0bXIkrk021a7lAO
# Yq9kqVDFv8K8O5WYvNcvbtAg3QW5JEaFnM3TMaOOSaWZMmIo7lw3e+B8rqknwmcS
# 66W2E0uayJXKqh/SXfS/xCwO2EzBT9Q1x0XiFR1LlEHQ0T/tfenBUlefIxfDZnEC
# AwEAAaOCAaQwggGgMA4GA1UdDwEB/wQEAwIHgDCBlAYIKwYBBQUHAQEEgYcwgYQw
# SAYIKwYBBQUHMAKGPGh0dHA6Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0
# L2dzY29kZXNpZ25zaGEyZzNvY3NwLmNydDA4BggrBgEFBQcwAYYsaHR0cDovL29j
# c3AyLmdsb2JhbHNpZ24uY29tL2dzY29kZXNpZ25zaGEyZzMwVgYDVR0gBE8wTTBB
# BgkrBgEEAaAyATIwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFsc2ln
# bi5jb20vcmVwb3NpdG9yeS8wCAYGZ4EMAQQBMAkGA1UdEwQCMAAwPwYDVR0fBDgw
# NjA0oDKgMIYuaHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9nc2NvZGVzaWduc2hh
# MmczLmNybDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUZLedJVdZSZd5
# lwNJFEgIc8KbEFEwHwYDVR0jBBgwFoAUDzrnrJSRdC2WAnODrZwuST8ZqlQwDQYJ
# KoZIhvcNAQELBQADggEBADYcz/+SCP59icPJK5w50yiTcoxnOtoA21GZDpt4GGVf
# RQJDWCDJMkU62xwu5HzqwimbwmBykrAf5Log1fLbggI83zIE4sMjkUe/BnnHpHgK
# LYv+3eLEwglMw/6Gmlq9IqNSD8YmTncGZFoFhrCrgAZUkA6RiVxuZrx2wiluueBI
# vfGs+tRA+7Tgx6Ed9kBybnc+xbAiTCNIcSo9OkPZfc3Q9saMgjIehBMXHLgMdrhv
# N5HXv/r4+aZ6asgv3ggArHrS1Pxp0f60hooVK4bA4Ph1td6YZ5lf8HA4uMmHvOjQ
# iNS0UjXqu5Vs6leIRM3pBjuX45xL6ydUsMlLhZQfanswggZqMIIFUqADAgECAhAD
# AZoCOv9YsWvW1ermF/BmMA0GCSqGSIb3DQEBBQUAMGIxCzAJBgNVBAYTAlVTMRUw
# EwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20x
# ITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTAeFw0xNDEwMjIwMDAw
# MDBaFw0yNDEwMjIwMDAwMDBaMEcxCzAJBgNVBAYTAlVTMREwDwYDVQQKEwhEaWdp
# Q2VydDElMCMGA1UEAxMcRGlnaUNlcnQgVGltZXN0YW1wIFJlc3BvbmRlcjCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKNkXfx8s+CCNeDg9sYq5kl1O8xu
# 4FOpnx9kWeZ8a39rjJ1V+JLjntVaY1sCSVDZg85vZu7dy4XpX6X51Id0iEQ7Gcnl
# 9ZGfxhQ5rCTqqEsskYnMXij0ZLZQt/USs3OWCmejvmGfrvP9Enh1DqZbFP1FI46G
# RFV9GIYFjFWHeUhG98oOjafeTl/iqLYtWQJhiGFyGGi5uHzu5uc0LzF3gTAfuzYB
# je8n4/ea8EwxZI3j6/oZh6h+z+yMDDZbesF6uHjHyQYuRhDIjegEYNu8c3T6Ttj+
# qkDxss5wRoPp2kChWTrZFQlXmVYwk/PJYczQCMxr7GJCkawCwO+k8IkRj3cCAwEA
# AaOCAzUwggMxMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB
# /wQMMAoGCCsGAQUFBwMIMIIBvwYDVR0gBIIBtjCCAbIwggGhBglghkgBhv1sBwEw
# ggGSMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMIIB
# ZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkA
# cwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUA
# cwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkA
# QwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkA
# aQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMA
# aAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIA
# ZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkA
# IAByAGUAZgBlAHIAZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTAfBgNVHSMEGDAWgBQV
# ABIrE5iymQftHt+ivlcNK2cCzTAdBgNVHQ4EFgQUYVpNJLZJMp1KKnkag0v0HonB
# yn0wfQYDVR0fBHYwdDA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEQ0EtMS5jcmwwOKA2oDSGMmh0dHA6Ly9jcmw0LmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3JsMHcGCCsGAQUFBwEBBGsw
# aTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUF
# BzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURDQS0xLmNydDANBgkqhkiG9w0BAQUFAAOCAQEAnSV+GzNNsiaBXJuGziMgD4CH
# 5Yj//7HUaiwx7ToXGXEXzakbvFoWOQCd42yE5FpA+94GAYw3+puxnSR+/iCkV61b
# t5qwYCbqaVchXTQvH3Gwg5QZBWs1kBCge5fH9j/n4hFBpr1i2fAnPTgdKG86Ugnw
# 7HBi02JLsOBzppLA044x2C/jbRcTBu7kA7YUq/OPQ6dxnSHdFMoVXZJB2vkPgdGZ
# dA0mxA5/G7X1oPHGdwYoFenYk+VVFvC7Cqsc21xIJ2bIo4sKHOWV2q7ELlmgYd3a
# 822iYemKC23sEhi991VUQAOSK2vCUcIKSK+w1G7g9BQKOhvjjz3Kr2qNe9zYRDCC
# Bs0wggW1oAMCAQICEAb9+QOWA63qAArrPye7uhswDQYJKoZIhvcNAQEFBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTA2MTExMDAwMDAwMFoXDTIxMTExMDAwMDAwMFowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6IItmfnKwkKVpYBzQHDSnlZUXKnE0kEG
# j8kz/E1FkVyBn+0snPgWWd+etSQVwpi5tHdJ3InECtqvy15r7a2wcTHrzzpADEZN
# k+yLejYIA6sMNP4YSYL+x8cxSIB8HqIPkg5QycaH6zY/2DDD/6b3+6LNb3Mj/qxW
# BZDwMiEWicZwiPkFl32jx0PdAug7Pe2xQaPtP77blUjE7h6z8rwMK5nQxl0SQoHh
# g26Ccz8mSxSQrllmCsSNvtLOBq6thG9IhJtPQLnxTPKvmPv2zkBdXPao8S+v7Iki
# 8msYZbHBc63X8djPHgp0XEK4aH631XcKJ1Z8D2KkPzIUYJX9BwSiCQIDAQABo4ID
# ejCCA3YwDgYDVR0PAQH/BAQDAgGGMDsGA1UdJQQ0MDIGCCsGAQUFBwMBBggrBgEF
# BQcDAgYIKwYBBQUHAwMGCCsGAQUFBwMEBggrBgEFBQcDCDCCAdIGA1UdIASCAckw
# ggHFMIIBtAYKYIZIAYb9bAABBDCCAaQwOgYIKwYBBQUHAgEWLmh0dHA6Ly93d3cu
# ZGlnaWNlcnQuY29tL3NzbC1jcHMtcmVwb3NpdG9yeS5odG0wggFkBggrBgEFBQcC
# AjCCAVYeggFSAEEAbgB5ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIA
# dABpAGYAaQBjAGEAdABlACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMA
# ZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAA
# QwBQAC8AQwBQAFMAIABhAG4AZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAA
# YQByAHQAeQAgAEEAZwByAGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0A
# aQB0ACAAbABpAGEAYgBpAGwAaQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMA
# bwByAHAAbwByAGEAdABlAGQAIABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUA
# cgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMVMBIGA1UdEwEB/wQIMAYBAf8CAQAweQYI
# KwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmww
# OqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcmwwHQYDVR0OBBYEFBUAEisTmLKZB+0e36K+Vw0rZwLNMB8GA1Ud
# IwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQBG
# UD7Jtygkpzgdtlspr1LPUukxR6tWXHvVDQtBs+/sdR90OPKyXGGinJXDUOSCuSPR
# ujqGcq04eKx1XRcXNHJHhZRW0eu7NoR3zCSl8wQZVann4+erYs37iy2QwsDStZS9
# Xk+xBdIOPRqpFFumhjFiqKgz5Js5p8T1zh14dpQlc+Qqq8+cdkvtX8JLFuRLcEwA
# iR78xXm8TBJX/l/hHrwCXaj++wc4Tw3GXZG5D2dFzdaD7eeSDY2xaYxP+1ngIw/S
# qq4AfO6cQg7PkdcntxbuD8O9fAqg7iwIVYUiuOsYGk38KiGtSTGDR5V3cdyxG0tL
# HBCcdxTBnU8vWpUIKRAmMYIEHjCCBBoCAQEwajBaMQswCQYDVQQGEwJCRTEZMBcG
# A1UEChMQR2xvYmFsU2lnbiBudi1zYTEwMC4GA1UEAxMnR2xvYmFsU2lnbiBDb2Rl
# U2lnbmluZyBDQSAtIFNIQTI1NiAtIEczAgxR+vwn6MLKvgDLtg0wCQYFKw4DAhoF
# AKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisG
# AQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcN
# AQkEMRYEFP2rodZS3AGpjTrL2GWSqSkDR+AZMA0GCSqGSIb3DQEBAQUABIIBAGeU
# lEnXiZTEQ/uZ9M9GnKCECK/lqTfpphxq6+4hVn3jhLqyrSSo/SmunwWa+E2pgJ98
# zndLupuF6u3IO2G9vaSWrrPQPhfF8FwUncxvXfMOLe5PDRMB1O/XSvCnCQv9cdo8
# uxZLli8TFJJ1jWxk2NXtXLg6L/QcVqaUVlNGr5XdXmIerh1iH8wdmDWuz/9cZC1S
# +6xvrlu+nJfKX9nYyY5M+NjglWIKPDbyYhs0ZCyozGaMc2fxgs4hpmiIRNrG6kM9
# q8LwvnX05ot3FuXYuh/d7lfa3+SogIPJB+zm1eI7Sw4z7x29biXjzsf4+BQkW4m7
# Plw5h5TzRaKLpYPQCT+hggIPMIICCwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBi
# MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3
# d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENB
# LTECEAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzEL
# BgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE5MDUyODEzMjEyNFowIwYJKoZI
# hvcNAQkEMRYEFKlyR4p6iRRBtBBbF1/afoJedSzJMA0GCSqGSIb3DQEBAQUABIIB
# AFu7tOUn0YVVTgXsYYnKXMzBbiFvUCZaWBU73hX8/HOB5nrSgZvYuy61qJE5sM8S
# Z4MxJpxtmHnkKkMLdLTT6cvSde81UsoV8FS4gP7H6YMf0Otmqpgt8ysPkvUwF7rW
# dB3ebOtNaz3RnMhRCv32F0PiTUBo1VhtPkYDAKGNoboqfJl/UrzLqpD6ahzDkvAw
# Hen+UUzovkCTqE10PD0EJE2tsFeUKt7iTmem6b+iuXzzpZL5R6Q9idLRP0Skofra
# kSWP/u/OpvET7zF7c4QuFsGjVjfnLrTNjQoGL96PVNM0n1JLF/OuT/Uy6SFzdUQL
# YCvEYjK0r4S6IaoLDPSztpY=
# SIG # End signature block
