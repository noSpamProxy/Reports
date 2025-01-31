<#
.SYNOPSIS
  Name: Get-NspLicenseReport.ps1
  Create NSP license report from the last 90 days.

.DESCRIPTION
  This script can be used to generate a license report either user-based, domain-based or both.
  It is possible to send the report via E-Mail to one or multiple recipients.
	This script uses the NoSpamProxy Powershell Cmdlets and an SQL query to generate the report files.
	The report will be generated always for the past 90 days.

.PARAMETER GDPR
	Generates a report with hashed names and e-mail addresses.    
    
.PARAMETER Minimal
	The output file will only contain the numeric number of used licenses.

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

.PARAMETER TenantId
	Default: 0
	Provide the tenant for which a license report should be generated.

.PARAMETER WithEncryptionDetails
	Enables a detailed report for the encryption module.
	The output will be splittet into SMIME/PGP encryption and decryption and PDF Encryption.    
    
.OUTPUTS
	Report is temporary stored under %TEMP% if the report is send via by E-Mail.
	If the parameter <NoMail> is used the files will be saved at the current location of the executing user.

.NOTES
  Version:        1.2.0
  Author:         Jan Jaeschke
  Creation Date:  2022-03-09
  Purpose/Change: added local password file usage
  
.LINK
  https://www.nospamproxy.de
  https://forum.nospamproxy.de
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
	.\Get-NspLicenseReport.ps1 -NoMail -SqlServer sql.example.com -SqlInstance NSPIntranetRole -SqlDatabase NSPIntranet -SqlCredential $Credentials
	This generates a user-based report. Therefore the script connects to the SQL Server "sql.example.com" and accesses the SQL instance "NSPIntranetRole" which contains the "NSPIntranet" database.
	The passed varaible "$Credentials" contains the desired user credentials. (e.x. $Credentials = Get-Credentials)

.EXAMPLE 
	.\Get-NspLicenseReport.ps1 -NoMail -SqlInstance ""
	Use the above instance name "" if you try to access the default SQL instance.
	If there is aconnection problem and the NSP configuration shows an empty instance for the intranet-role under "Configuration -> NoSpamProxy components" than this instance example should work.

.EXAMPLE
	.\Get-NspLicenseReport.ps1 -NoMail -SqlInstance "" -TenantId 42
	NoSpamProxy v14 provides a provider mode which allows the usage of multiple tenants.
	To generate a report for a specific tenant it is necessary to provide the desired tenant id.

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
	[Parameter(Mandatory=$false)][ValidatePattern("^([a-zA-Z0-9\s.!£#$%&'^_`{}~-]+)?<?[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*>?$")][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	# change report E-Mail subject
	[Parameter(Mandatory=$false)][string] $ReportSubject = "NSP License Report",
	[Parameter(Mandatory=$false)][ValidateSet('user-based','domain-based','both')][string] $ReportType = "user-based",
	# set used SMTP host for sending report E-Mail only a valid  IP address or FQDN is allowed
	[Parameter(Mandatory=$false)][ValidatePattern("^(((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))|(((?!-)[a-zA-Z0-9-]{0,62}[a-zA-Z0-9]\.?)+[a-zA-Z]{2,63}))$")][string] $SmtpHost,
	# sql credentials
	[Parameter(Mandatory=$false)][pscredential] $SqlCredential,
    # sql username only used if password is saved in an encrypted binary file
    [Parameter(Mandatory = $false)][string] $SqlUsername,
    # locationnof the encrypted binary file for saving a sql password
    [Parameter(Mandatory=$false)]
        [string] $SqlPasswordFileLocation = "$PSScriptRoot\NspReadSqlPass.bin",
	# database name
	[Parameter(Mandatory=$false)][string] $SqlDatabase = "NoSpamProxyAddressSynchronization",
	# sql server instance
	[Parameter(Mandatory=$false)][string] $SqlInstance = "NoSpamProxy",
	# sql server
	[Parameter(Mandatory=$false)][string] $SqlServer = "(local)",
	# generate detailed report including which encryption features are used
	[Parameter(Mandatory=$false)][switch] $WithEncryptionDetails,
	# hash all name and e-mail addresses
	[Parameter(Mandatory=$false)][switch] $GDPRconform,
	# ouput the lic count number only
	[Parameter(Mandatory=$false)][switch] $Minimal,
    # tenant id for generation 
	[Parameter(Mandatory=$false)][int] $TenantId = 0,
    # set SqlUser password
    [Parameter(Mandatory=$false)]
        [switch] $SetLoginPassword,
	# enable the output for managed certificates (only available in some environments)
	[Parameter(Mandatory=$false)][switch] $WithManagedCertificates
)

#-------------------Functions----------------------
# create an encrypted binary file which contains the password for the desired SQL user
function Set-loginPass {
    # Imports Security library for encryption
    Add-Type -AssemblyName System.Security
    $sqlPass = Read-Host -Promp 'Input your user password'
    $passFileLocation = "$SqlPasswordFileLocation"
    $inBytes = [System.Text.Encoding]::Unicode.GetBytes($sqlPass)
    $protected = [System.Security.Cryptography.ProtectedData]::Protect($inBytes, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    [System.IO.File]::WriteAllBytes($passFileLocation, $protected)
}
# return the SQL user password from the encrypted binary file
# if file does not exists give a hint and ask for manual input
function Get-loginPass {
    # Imports Security library for encryption
    Add-Type -AssemblyName System.Security
    Add-Type -AssemblyName System.Text.Encoding
    $passFileLocation = "$SqlPasswordFileLocation"
    if (Test-Path $passFileLocation) {
        try {
            $protected = [System.IO.File]::ReadAllBytes($passFileLocation)
            $rawKey = [System.Security.Cryptography.ProtectedData]::Unprotect($protected, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
            return [System.Text.Encoding]::Unicode.GetString($rawKey)
        }
        catch {
            Write-Host $_.Exception | format-list -force
        }
    }
    else {
        Write-Host "No Password file found! Please run '$($PSCommandPath) -SetLoginPassword' for saving your password encrypted."
        $loginPass = Read-Host -Promp 'Input your user password'
        return $loginPass
    }
}
# create sha265 hash
function hashValue($userArray){
	$hashedUsers=@()
	$hasher = [System.Security.Cryptography.HashAlgorithm]::Create('sha256')
	foreach ($user in $userArray){
		$hash = $hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($user))
		$hashString = [System.BitConverter]::ToString($hash)
		$hashedUsers += $hashString.Replace('-', '')
	}
	return $hashedUsers
}
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
	if ($SqlCredential) {
		$networkCredential = $SqlCredential.GetNetworkCredential()
		$connectionString += "uid=" + $networkCredential.UserName + ";pwd=" + $networkCredential.Password + ";"
    } elseif ($SqlUsername) {
        $password = (convertto-securestring -string (Get-loginPass) -asplaintext -force)
        $Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $SqlUsername, $password
        $networkCredential = $Credential.GetNetworkCredential()
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
			$command.CommandText = (Get-Content "$PSScriptRoot\$queryName.sql") -f $TenantId
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

function generateReport($reportName, $varTable, $WithEncryptionDetails){
	$encoding = New-Object System.Text.UTF8Encoding $false
	$stream = New-Object System.IO.StreamWriter $reportName, $false, $encoding
	$stream.WriteLine("Protection: $protectionLicCount User")
	$stream.WriteLine("-----------------------------")
	if(!$Minimal){
		$varTable.protectionUsers | ForEach-Object{$stream.WriteLine($_)}
	}
	$stream.WriteLine("`r`n`r`nEncryption: $encryptionLicCount User")
	$stream.WriteLine("-----------------------------")
	if(!$Minimal){
		$varTable.encryptionUsers | ForEach-Object{$stream.WriteLine($_)}
	}
	$stream.WriteLine("`r`n`r`nLargeFiles: $largeFilesLicCount User")
	$stream.WriteLine("-----------------------------")
	if(!$Minimal){
		$varTable.largeFilesUsers | ForEach-Object{$stream.WriteLine($_)}
	}
	$stream.WriteLine("`r`n`r`nDisclaimer: $disclaimerLicCount User")
	$stream.WriteLine("-----------------------------")
	if(!$Minimal){
		$varTable.disclaimerUsers | ForEach-Object{$stream.WriteLine($_)}
	}
	$stream.WriteLine("`r`n`r`nSandBox: $sandBoxLicCount User")
	$stream.WriteLine("-----------------------------")
	if(!$Minimal){
		$varTable.sandBoxUsers | ForEach-Object{$stream.WriteLine($_)}
	}
	# print detailed information if parameter flag is used
	if($WithEncryptionDetails){
		$stream.WriteLine("`r`n`r`n`r`nEncryption details")
		$stream.WriteLine("-----------------------------")
		$stream.WriteLine("Please be warned that users can be displayed multiple times.")
		$stream.WriteLine("In this case the number count is not affected.")
        $stream.WriteLine("-----------------------------")
        $stream.WriteLine("`r`nPDF Mail: $($varTable.pdfMailsSentCount) User")
        $stream.WriteLine("-----------------------------")
        $varTable.pdfMailsSentUsers | ForEach-Object { $stream.WriteLine($_) }
        $stream.WriteLine("`r`n`r`nSMIME signature: $($varTable.sMimeMailsSignedCount) User")
        $stream.WriteLine("-----------------------------")
        $varTable.sMimeMailsSignedUsers | ForEach-Object { $stream.WriteLine($_) }
        $stream.WriteLine("`r`n`r`nSMIME encryption: $($varTable.sMimeMailsEncryptedCount) User")
        $stream.WriteLine("-----------------------------")
        $varTable.sMimeMailsEncryptedUsers | ForEach-Object { $stream.WriteLine($_) }
        $stream.WriteLine("`r`n`r`nSMIME validation: $($varTable.sMimeMailsValidatedCount) User")
        $stream.WriteLine("-----------------------------")
        $varTable.sMimeMailsValidatedUsers | ForEach-Object { $stream.WriteLine($_) }
        $stream.WriteLine("`r`n`r`nSMIME decrypted: $($varTable.sMimeMailsDecryptedCount) User")
        $stream.WriteLine("-----------------------------")
        $varTable.sMimeMailsDecryptedUsers | ForEach-Object { $stream.WriteLine($_) }
        $stream.WriteLine("`r`n`r`PGP signature: $($pgpMailsSignedCount) User")
        $stream.WriteLine("-----------------------------")
        $varTable.pgpMailsSignedUsers | ForEach-Object { $stream.WriteLine($_) }
        $stream.WriteLine("`r`n`r`PGP encryption: $($pgpMailsEncryptedCount) User")
        $stream.WriteLine("-----------------------------")
        $varTable.pgpMailsEncryptedUsers | ForEach-Object { $stream.WriteLine($_) }
	} if ($WithManagedCertificates) {
		$stream.WriteLine("`r`n`r`nManagedCertificates: $($subjects.Count)")
		$stream.WriteLine("-----------------------------")
		$stream.WriteLine($($subjects | Out-String))
	}
	$stream.Close()
}

# generates user-based report
function userBased($licUsage){
	$varTable = @{}
	# the @ sign in each .Count line is needed in the case there is only 1 item
	# in this case PS return this single item instead of an array which breaks .Count
	# get license count per feature
	$protectionLicCount = @($licUsage | Where-Object{$_.Protection -eq 1}).Count
	$encryptionLicCount = @($licUsage | Where-Object{$_.Encryption -eq 1}).Count
	$largeFilesLicCount = @($licUsage | Where-Object{$_.LargeFiles -eq 1}).Count
	$disclaimerLicCount = @($licUsage | Where-Object{$_.Disclaimer -eq 1}).Count
	$sandBoxLicCount = @($licUsage | Where-Object{$_.FilesUploadedToSandbox -eq 1}).Count
	# add lic count to hashtable
	$varTable.Add("protectionLicCount", $protectionLicCount)
	$varTable.Add("encryptionLicCount", $encryptionLicCount)
	$varTable.Add("largeFilesLicCount", $largeFilesLicCount)
	$varTable.Add("disclaimerLicCount", $disclaimerLicCount)
	$varTable.Add("sandBoxLicCount", $sandBoxLicCount)
	# if script is called by Get-NspLicenseReportForMsp, set global variables for simple data exchange
	if ((Get-Variable MyInvocation -Scope 1).Value.MyCommand.CommandType -eq "ExternalScript") {
		Write-Host "Remote invoke"
		$global:protectionCount += $protectionLicCount
		$global:encryptionCount += $encryptionLicCount
		$global:largeFilesCount += $largeFilesLicCount
		$global:disclaimerCount += $disclaimerLicCount
		$global:sandBoxCount += $sandBoxLicCount
	}
	# get users per feature
	$protectionUsers = ($licUsage | Where-Object{$_.Protection -eq 1}).DisplayName
	$encryptionUsers = ($licUsage | Where-Object{$_.Encryption -eq 1}).DisplayName
	$largeFilesUsers = ($licUsage | Where-Object{$_.LargeFiles -eq 1}).DisplayName
	$disclaimerUsers = ($licUsage | Where-Object{$_.Disclaimer -eq 1}).DisplayName
	$sandBoxUsers = ($licUsage | Where-Object{$_.FilesUploadedToSandbox -eq 1}).DisplayName
	if ($GDPRconform){
		$userReport = "$reportFile" + "_per_user_gdpr.txt"
		# add hashed users to hastable
		$hashedProtectionUsers = hashValue $protectionUsers
		$hashedEncryptionUsers= hashValue $encryptionUsers
		$hashedLargeFilesUsers = hashValue $largeFilesUsers
		$hashedDisclaimerUsers = hashValue $disclaimerUsers
		$hashedSandBoxUsers = hashValue $sandBoxUsers
		$varTable.Add("protectionUsers", $hashedProtectionUsers)
		$varTable.Add("encryptionUsers", $hashedEncryptionUsers)
		$varTable.Add("largeFilesUsers", $hashedLargeFilesUsers)
		$varTable.Add("disclaimerUsers", $hashedDisclaimerUsers)
		$varTable.Add("sandBoxUsers", $hashedSandBoxUsers)
	}else{
		$userReport = "$reportFile" + "_per_user.txt"
		# add unhashed users to hastable
		$varTable.Add("protectionUsers", $protectionUsers)
		$varTable.Add("encryptionUsers", $encryptionUsers)
		$varTable.Add("largeFilesUsers", $largeFilesUsers)
		$varTable.Add("disclaimerUsers", $disclaimerUsers)
		$varTable.Add("sandBoxUsers", $sandBoxUsers)
	}
	# parse detailed information if parameter flag is used
	if ($WithEncryptionDetails){
		$pdfMailsSentCount = @($licUsage | Where-Object{$_.PdfMailsSent -eq 1}).Count
		$sMimeMailsSignedCount = @($licUsage | Where-Object{$_.SMimeMailsSigned -eq 1}).Count
		$sMimeMailsEncryptedCount = @($licUsage | Where-Object{$_.SMimeMailsEncrypted -eq 1}).Count
		$sMimeMailsValidatedCount = ($licUsage | Where-Object{$_.SMimeMailsValidated -eq 1}).Count
		$sMimeMailsDecryptedCount = ($licUsage | Where-Object{$_.SMimeMailsDecrypted -eq 1}).Count
		$pgpMailsSignedCount = @($licUsage | Where-Object{$_.PgpMailsSigned -eq 1}).Count
		$pgpMailsEncryptedCount = @($licUsage | Where-Object{$_.PgpMailsEncrypted -eq 1}).Count

		$varTable.Add("pdfMailsSentCount", $pdfMailsSentCount)
		$varTable.Add("sMimeMailsSignedCount", $sMimeMailsSignedCount)
		$varTable.Add("sMimeMailsEncryptedCount", $sMimeMailsEncryptedCount)
		$varTable.Add("sMimeMailsValidatedCount", $sMimeMailsValidatedCount)
		$varTable.Add("sMimeMailsDecryptedCount", $sMimeMailsDecryptedCount)
		$varTable.Add("pgpMailsSignedCount", $pgpMailsSignedCount)
		$varTable.Add("pgpMailsEncryptedCount", $pgpMailsEncryptedCount)

		$pdfMailsSentUsers = ($licUsage | Where-Object{$_.PdfMailsSent -eq 1}).DisplayName
		$sMimeMailsSignedUsers = ($licUsage | Where-Object{$_.SMimeMailsSigned -eq 1}).DisplayName
		$sMimeMailsEncryptedUsers = ($licUsage | Where-Object{$_.SMimeMailsEncrypted -eq 1}).DisplayName
		$sMimeMailsValidatedUsers = ($licUsage | Where-Object{$_.SMimeMailsValidated -eq 1}).DisplayName
		$sMimeMailsDecryptedUsers = ($licUsage | Where-Object{$_.SMimeMailsDecrypted -eq 1}).DisplayName
		$pgpMailsSignedUsers = ($licUsage | Where-Object{$_.PgpMailsSigned -eq 1}).DisplayName
		$pgpMailsEncryptedUsers = ($licUsage | Where-Object{$_.PgpMailsEncrypted -eq 1}).DisplayName

		if ($GDPRconform){
			$hashedPdfMailsSentUsers = hashValue $pdfMailsSentUsers
			$varTable.Add("pdfMailsSentUsers", $hashedPdfMailsSentUsers)
			$hashedSmimeMailsSignedUsers = hashValue $sMimeMailsSignedUsers
			$varTable.Add("sMimeMailsSignedUsers", $hashedSmimeMailsSignedUsers)
			$hashedSmimeMailsEncryptedUsers = hashValue $sMimeMailsEncryptedUsers
			$varTable.Add("sMimeMailsEncryptedUsers", $hashedSmimeMailsEncryptedUsers)
			$hashedSmimeMailsValidatedUsers = hashValue $sMimeMailsValidatedUsers
			$varTable.Add("sMimeMailsValidatedUsers", $hashedSmimeMailsValidatedUsers)
			$hashedSmimeMailsDecryptedUsers = hashValue $sMimeMailsDecryptedUsers
			$varTable.Add("sMimeMailsDecryptedUsers", $hashedSmimeMailsDecryptedUsers)
			$hashedPgpMailsSignedUsers = hashValue $pgpMailsSignedUsers
			$varTable.Add("pgpMailsSignedUsers", $hashedPgpMailsSignedUsers)
			$hashedPgpMailsEncryptedUsers = hashValue $pgpMailsEncryptedUsers
			$varTable.Add("pgpMailsEncryptedUsers", $hashedPgpMailsEncryptedUsers)
		}else{
			$varTable.Add("pdfMailsSentUsers", $pdfMailsSentUsers)
			$varTable.Add("sMimeMailsSignedUsers", $sMimeMailsSignedUsers)
			$varTable.Add("sMimeMailsEncryptedUsers", $sMimeMailsEncryptedUsers)
			$varTable.Add("sMimeMailsValidatedUsers", $sMimeMailsValidatedUsers)
			$varTable.Add("sMimeMailsDecryptedUsers", $sMimeMailsDecryptedUsers)
			$varTable.Add("pgpMailsSignedUsers", $pgpMailsSignedUsers)
			$varTable.Add("pgpMailsEncryptedUsers", $pgpMailsEncryptedUsers)
		}
		generateReport $userReport $varTable $true
	}else{
		generateReport $userReport $varTable $false
	}
	return $userReport
}

# generates domain-based report
function domainBased($licUsage){
	$ownDomains = ($licUsage.Domain | Get-Unique)
	$domainReportFiles = @()
	foreach($domain in $ownDomains){
		$varTable = @{}
		# the @ sign in each .Count line is needed in the case there is only 1 item
		# in this case PS return this single item instead of an array which breaks .Count
		# get license count per feature
		$protectionLicCount = @($licUsage | Where-Object {$_.Protection -eq 1 -and $_.Domain -eq $domain}).Count
		$encryptionLicCount = @($licUsage | Where-Object {$_.Encryption -eq 1 -and $_.Domain -eq $domain}).Count
		$largeFilesLicCount = @($licUsage | Where-Object {$_.LargeFiles -eq 1 -and $_.Domain -eq $domain}).Count
		$disclaimerLicCount = @($licUsage | Where-Object {$_.Disclaimer -eq 1 -and $_.Domain -eq $domain}).Count
		$sandBoxLicCount = @($licUsage | Where-Object {$_.FilesUploadedToSandbox -eq 1 -and $_.Domain -eq $domain}).Count
		# add lic count to hashtable
		$varTable.Add("protectionLicCount", $protectionLicCount)
		$varTable.Add("encryptionLicCount", $encryptionLicCount)
		$varTable.Add("largeFilesLicCount", $largeFilesLicCount)
		$varTable.Add("disclaimerLicCount", $disclaimerLicCount)
		$varTable.Add("sandBoxLicCount", $sandBoxLicCount)
		# if script is called by Get-NspLicenseReportForMsp, set global variables for simple data exchange
		if ((Get-Variable MyInvocation -Scope 1).Value.MyCommand.CommandType -eq "ExternalScript") {
			$global:protectionCount += $protectionLicCount
			$global:encryptionCount += $encryptionLicCount
			$global:largeFilesCount += $largeFilesLicCount
			$global:disclaimerCount += $disclaimerLicCount
			$global:sandBoxCount += $sandBoxLicCount
		}
		# get users per feature
		$protectionUsers = ($licUsage | Where-Object{$_.Protection -eq 1 -and $_.Domain -eq $domain}).DisplayName
		$encryptionUsers = ($licUsage | Where-Object{$_.Encryption -eq 1 -and $_.Domain -eq $domain}).DisplayName
		$largeFilesUsers = ($licUsage | Where-Object{$_.LargeFiles -eq 1 -and $_.Domain -eq $domain}).DisplayName
		$disclaimerUsers = ($licUsage | Where-Object{$_.Disclaimer -eq 1 -and $_.Domain -eq $domain}).DisplayName
		$sandBoxUsers = ($licUsage | Where-Object{$_.FilesUploadedToSandbox -eq 1 -and $_.Domain -eq $domain}).DisplayName
		if ($GDPRconform){
			$domainReport = "$reportFile" + "_" + "$domain" + "_gdpr.txt"
			# add hashed users to hastable
			$hashedProtectionUsers = hashValue $protectionUsers
			$hashedEncryptionUsers= hashValue $encryptionUsers
			$hashedLargeFilesUsers = hashValue $largeFilesUsers
			$hashedDisclaimerUsers = hashValue $disclaimerUsers
			$hashedSandBoxUsers = hashValue $sandBoxUsers
			$varTable.Add("protectionUsers", $hashedProtectionUsers)
			$varTable.Add("encryptionUsers", $hashedEncryptionUsers)
			$varTable.Add("largeFilesUsers", $hashedLargeFilesUsers)
			$varTable.Add("disclaimerUsers", $hashedDisclaimerUsers)
			$varTable.Add("sandBoxUsers", $hashedSandBoxUsers)
		}else{
			$domainReport = "$reportFile" + "_" + "$domain" +".txt"
			# add unhashed users to hastable
			$varTable.Add("protectionUsers", $protectionUsers)
			$varTable.Add("encryptionUsers", $encryptionUsers)
			$varTable.Add("largeFilesUsers", $largeFilesUsers)
			$varTable.Add("disclaimerUsers", $disclaimerUsers)
			$varTable.Add("sandBoxUsers", $sandBoxUsers)
		}
		# parse detailed information if parameter flag is used
		if ($WithEncryptionDetails){
			$pdfMailsSentCount = @($licUsage | Where-Object{$_.PdfMailsSent -eq 1 -and $_.Domain -eq $domain}).Count
			$sMimeMailsSignedCount = @($licUsage | Where-Object{$_.SMimeMailsSigned -eq 1 -and $_.Domain -eq $domain}).Count
			$sMimeMailsEncryptedCount = @($licUsage | Where-Object{$_.SMimeMailsEncrypted -eq 1 -and $_.Domain -eq $domain}).Count
			$sMimeMailsValidatedCount = @($licUsage | Where-Object{$_.SMimeMailsValidated -eq 1 -and $_.Domain -eq $domain}).Count
			$sMimeMailsDecryptedCount = @($licUsage | Where-Object{$_.SMimeMailsDecrypted -eq 1 -and $_.Domain -eq $domain}).Count
			$pgpMailsSignedCount = @($licUsage | Where-Object{$_.PgpMailsSigned -eq 1 -and $_.Domain -eq $domain}).Count
			$pgpMailsEncryptedCount = @($licUsage | Where-Object{$_.PgpMailsEncrypted -eq 1 -and $_.Domain -eq $domain}).Count
	
			$varTable.Add("pdfMailsSentCount", $pdfMailsSentCount)
			$varTable.Add("sMimeMailsSignedCount", $sMimeMailsSignedCount)
			$varTable.Add("sMimeMailsEncryptedCount", $sMimeMailsEncryptedCount)
			$varTable.Add("sMimeMailsValidatedCount", $sMimeMailsValidatedCount)
			$varTable.Add("sMimeMailsDecryptedCount", $sMimeMailsDecryptedCount)
			$varTable.Add("pgpMailsSignedCount", $pgpMailsSignedCount)
			$varTable.Add("pgpMailsEncryptedCount", $pgpMailsEncryptedCount)

			$pdfMailsSentUsers = ($licUsage | Where-Object{$_.PdfMailsSent -eq 1 -and $_.Domain -eq $domain}).DisplayName
			$sMimeMailsSignedUsers = ($licUsage | Where-Object{$_.SMimeMailsSigned -eq 1 -and $_.Domain -eq $domain}).DisplayName
			$sMimeMailsEncryptedUsers = ($licUsage | Where-Object{$_.SMimeMailsEncrypted -eq 1 -and $_.Domain -eq $domain}).DisplayName
			$sMimeMailsValidatedUsers = ($licUsage | Where-Object{$_.SMimeMailsValidated -eq 1 -and $_.Domain -eq $domain}).DisplayName
			$sMimeMailsDecryptedUsers = ($licUsage | Where-Object{$_.SMimeMailsDecrypted -eq 1 -and $_.Domain -eq $domain}).DisplayName
			$pgpMailsSignedUsers = ($licUsage | Where-Object{$_.PgpMailsSigned -eq 1 -and $_.Domain -eq $domain}).DisplayName
			$pgpMailsEncryptedUsers = ($licUsage | Where-Object{$_.PgpMailsEncrypted -eq 1 -and $_.Domain -eq $domain}).DisplayName

			if ($GDPRconform){
				$hashedPdfMailsSentUsers = hashValue $pdfMailsSentUsers
				$varTable.Add("pdfMailsSentUsers", $hashedPdfMailsSentUsers)
				$hashedSmimeMailsSignedUsers = hashValue $sMimeMailsSignedUsers
				$varTable.Add("sMimeMailsSignedUsers", $hashedSmimeMailsSignedUsers)
				$hashedSmimeMailsEncryptedUsers = hashValue $sMimeMailsEncryptedUsers
				$varTable.Add("sMimeMailsEncryptedUsers", $hashedSmimeMailsEncryptedUsers)
				$hashedSmimeMailsValidatedUsers = hashValue $sMimeMailsValidatedUsers
				$varTable.Add("sMimeMailsValidatedUsers", $hashedSmimeMailsValidatedUsers)
				$hashedSmimeMailsDecryptedUsers = hashValue $sMimeMailsDecryptedUsers
				$varTable.Add("sMimeMailsDecryptedUsers", $hashedSmimeMailsDecryptedUsers)
				$hashedPgpMailsSignedUsers = hashValue $pgpMailsSignedUsers
				$varTable.Add("pgpMailsSignedUsers", $hashedPgpMailsSignedUsers)
				$hashedPgpMailsEncryptedUsers = hashValue $pgpMailsEncryptedUsers
				$varTable.Add("pgpMailsEncryptedUsers", $hashedPgpMailsEncryptedUsers)
			}else{
				$varTable.Add("pdfMailsSentUsers", $pdfMailsSentUsers)
				$varTable.Add("sMimeMailsSignedUsers", $sMimeMailsSignedUsers)
				$varTable.Add("sMimeMailsEncryptedUsers", $sMimeMailsEncryptedUsers)
				$varTable.Add("sMimeMailsValidatedUsers", $sMimeMailsValidatedUsers)
				$varTable.Add("sMimeMailsDecryptedUsers", $sMimeMailsDecryptedUsers)
				$varTable.Add("pgpMailsSignedUsers", $pgpMailsSignedUsers)
				$varTable.Add("pgpMailsEncryptedUsers", $pgpMailsEncryptedUsers)
			}
			generateReport $domainReport $varTable $true
		}else{
			generateReport $domainReport $varTable $false
		}
		$domainReportFiles = $domainReportFiles + $domainReport
	}
	return $domainReportFiles
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
$reportFile =  "$reportFilePath" + "\" + "$reportFileDate" + "_" + "$ReportFileName" + "_" + "$TenantId"

#--------------------Main-----------------------
# create password file for login
if($SetLoginPassword){
    Set-loginPass
    EXIT
}
$databaseVersion = [Version] (Invoke-SqlQuery "SELECT value FROM sys.fn_listextendedproperty ('AddressSynchronizationDBVersion', null, null, null, null, null, default)" -isInlineQuery $true -isSingleResult $true)

if ($databaseVersion -gt ([Version] "14.0.100")) {
	$queryVersion = "v14"
	Write-Host "You are running NoSpamProxy in v14."
	Write-Host "If you enabled the provider mode make sure to pass the desired tenant id using the '-TenantId' parameter."
	if ($WithManagedCertificates) {
		$managedCertificatesReport = "$reportFile" + "_managedCertificates.txt"
		$managedCertificates = Invoke-SqlQuery "ManagedCertificates"
		if ($null -eq $managedCertificates.Subject ) {
			Write-Host "No managed certificates are licensed or requested."
			if ((Get-Variable MyInvocation -Scope 1).Value.MyCommand.CommandType -eq "ExternalScript") {
				$global:managedCertificateCount += 0
			}
		} else {
			if ($GDPRconform) {
				$subjects = $managedCertificates.Subject | ForEach-Object {hashValue($_.Substring(2, $($_.IndexOf(',')-2)))}
			} else {
				$subjects = $managedCertificates.Subject | ForEach-Object {$_.Substring(2, $($_.IndexOf(',')-2))}
			}
			if ((Get-Variable MyInvocation -Scope 1).Value.MyCommand.CommandType -eq "ExternalScript") {
				$global:managedCertificateCount += $($managedCertificates.Subject).Count
			}
		}
	}
}
else {
	$queryVersion = "v13"
}

if($ReportType){
	switch($ReportType){
		'user-based'{
			$licUsageUsers = Invoke-SqlQuery "LicenseUsageUsers_$queryVersion"
			$userReport = userBased $licUsageUsers
			# send mail if <NoMail> switch is not used and delete temp report file
			if (!$NoMail){
				sendMail $ReportRecipient $ReportRecipientCSV $userReport
				Remove-Item $userReport
			}
		}
		'domain-based'{
		$licUsageDomains = Invoke-SqlQuery "LicenseUsageDomains_$queryVersion"
			$domainReport = domainBased $licUsageDomains
			if (!$NoMail){
				sendMail $ReportRecipient $ReportRecipientCSV $domainReport
				Remove-Item $domainReport
			}
		}
		'both'{
			$licUsageUsers = Invoke-SqlQuery "LicenseUsageUsers_$queryVersion"
			$licUsageDomains = Invoke-SqlQuery "LicenseUsageDomains_$queryVersion"
			$userReport = userBased $licUsageUsers
			$domainReport = domainBased $licUsageDomains
			if (!$NoMail){
				$reportFileList = $domainReport
				$reportFileList += $userReport
				sendMail $ReportRecipient $ReportRecipientCSV $reportFileList
				Remove-Item $userReport
				Remove-Item $domainReport
			}
		}
	}
}
Write-Host "Skript durchgelaufen"