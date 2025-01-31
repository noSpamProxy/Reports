<#
.SYNOPSIS
  Name: Get-NspLicenseReportForMspForMsp.ps1
  Create NSP license report from the last 90 days with an overall license count.

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
  Version:        1.0.0
  Author:         Jan Jaeschke
  Creation Date:  2024-10-15
  Purpose/Change: initial wrapper script for Get-NspLivenseReport
  
.LINK
  https://www.nospamproxy.de
  https://forum.nospamproxy.de
  https://github.com/noSpamProxy

.EXAMPLE
  .\Get-NspLicenseReportForMsp.ps1 -ReportRecipient alice@example.com -ReportSender "NoSpamProxy Report Sender <nospamproxy@example.com>" -ReportSubject "Example Report" -SmtpHost mail.example.com
  
.EXAMPLE
  .\Get-NspLicenseReportForMsp.ps1 -ReportRecipient alice@example.com -ReportSender "NoSpamProxy Report Sender <nospamproxy@example.com>" -ReportSubject "Example Report" -SmtpHost mail.example.com -ReportRecipientCSV "C:\Users\example\Documents\email-recipient.csv"
  The CSV have to contain the header "Email" else the mail addresses cannot be read from the file. 
  E.g: email-recipient.csv
  User,Email
  user1,user1@example.com
	user2,user2@example.com
	The "User" header is not necessary.

.EXAMPLE
	.\Get-NspLicenseReportForMsp.ps1 -NoMail -ReportType both
	Generates a user-based and a domain-based report which are saved at the current location of execution, here: ".\"

.EXAMPLE
	.\Get-NspLicenseReportForMsp.ps1 -NoMail -SqlServer sql.example.com -SqlInstance NSPIntranetRole -SqlDatabase NSPIntranet -SqlCredential $Credentials
	This generates a user-based report. Therefore the script connects to the SQL Server "sql.example.com" and accesses the SQL instance "NSPIntranetRole" which contains the "NSPIntranet" database.
	The passed varaible "$Credentials" contains the desired user credentials. (e.x. $Credentials = Get-Credentials)

.EXAMPLE 
	.\Get-NspLicenseReportForMsp.ps1 -NoMail -SqlInstance ""
	Use the above instance name "" if you try to access the default SQL instance.
	If there is aconnection problem and the NSP configuration shows an empty instance for the intranet-role under "Configuration -> NoSpamProxy components" than this instance example should work.

.EXAMPLE
	.\Get-NspLicenseReportForMsp.ps1 -NoMail -SqlInstance "" -TenantId 42
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
	if ($(Get-NspIsProviderModeEnabled) -ne $true) {
		Write-Host "This script is only for provider enabled environments. Please use Get-NspLicenseReport.ps1 directly."
		EXIT
	}
} else {
		Write-Host "This script is only for provider enabled environments. Please use Get-NspLicenseReport.ps1 directly."
		EXIT
}

if ($PSBoundParameters.Reporttype -eq "both") {
	Write-Host "The overall license count for report type `'both`' will be incorrect. Use this mode only for report generation over all tenants."
}

$tenants = Get-Nsptenant
$global:protectionCount = 0 
$global:encryptionCount = 0 
$global:largeFilesCount = 0 
$global:disclaimerCount = 0 
$global:sandBoxCount = 0
$global:managedCertificateCount = 0
foreach ($tenant in $tenants) {
	#call script with tenan id
	$reportFileNameOverride = $ReportFileName + "$($tenant.Name)"
	$PSBoundParameters.ReportFileName = $reportFileNameOverride
	$PSBoundParameters.TenantId = $($tenant.Id)
	$PSBoundParameters.NoMail = $true
	& $PSSCriptroot\Get-NspLicenseReport.ps1 @PSBoundParameters
}

Write-Host "protection: $($global:protectionCount)"
Write-Host "encryption: $($global:encryptionCount)"
Write-Host "largefiles: $($global:largeFilesCount)"
Write-Host "disclaimer: $($global:disclaimerCount)"
Write-Host "sandbox: $($global:sandBoxCount)"
if ($PSBoundParameters.ManagedCertificates.IsPresent -eq $true) {
	Write-Host "managedcertificates: $($global:managedCertificateCount)"
}