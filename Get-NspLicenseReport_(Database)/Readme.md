
# Get-NspLicenseReport.ps1

Generates a report for the usage of each module.
It is possible to generate a user-based report which contains the following information:

 - Total Users who used the module
 - List of each user who used the module
 
Moreover it is possible to generate a domain-based report which contains the same information like above but separated for each domain. 

## Usage 

```ps
Get-NspLicenseReport.ps1 [-NoMail] [[-ReportFileName] <String>] [[-ReportRecipient] <String[]>] [[-ReportRecipientCSV] <String>][[-ReportSender] <String>] [[-ReportSubject] <String>] [[-ReportType] user-based | domain-based | both] [[-SmtpHost] <String>] [[-SqlCredential] <PSCredential>] [[-SqlDatabase] <String>] [[-SqlInstance]<String>] [[-SqlServer] <String>] [<CommonParameters>]
```

## Parameters
### NoMail
Does not generate an E-Mail and saves the generated reports to the current execution location.  
Ideal for testing or manual script usage.

### ReportFileName
**Default:** License_Report  
Define a part of the complete file name.
 
**user-based:**  
    C:\Users\example\Documents\2019-05-27_License_Report_per_user.txt  
**domain-based:**  
    C:\Users\example\Documents\2019-05-27_License_Report_example.com.txt  
	
### ReportRecipient
Specifies the E-Mail recipient. It is possible to pass a comma separated list to address multiple recipients.  
E.g.: alice@example.com, bob@example.com

### ReportRecipientCSV
Set a filepath to an CSV file containing a list of report E-Mail recipient. Be aware about the needed CSV format, please watch the provided example.

### ReportSender
**Default:** NoSpamProxy Report Sender <nospamproxy@example.com>
Sets the report E-Mail sender address.
  
### ReportSubject
**Default:** NSP License Report  
Sets the report E-Mail subject.
	
### ReportType
**Default:** user-based  
Sets the type of generated report.    
Possible values are: user-based, domain-based, both

### SmtpHost
Specifies the SMTP host which should be used to send the report E-Mail.  
It is possible to use a FQDN or IP address.
	
### SqlCredential
Sets custom credentials for database access.  
By default the authentication is done using current users credentials from memory.

### SqlDatabase
**Default:** NoSpamProxyAddressSynchronization  
Sets a custom SQl database name which should be accessed. The required database is the one from the intranet-role.

### SqlInstance
**Default:** NoSpamProxy  
Sets a custom SQL instance name which should be accessed. The required instance must contain the intranet-role database.

### SqlServer
**Default:** (local)  
Sets a custom SQL server which must contains the instance and the database of the intranet-role.

## Outputs
Report is temporary stored under %TEMP% if the report is send via by E-Mail.
If the parameter <NoMail> is used the files will be saved at the current location of the executing user.

## Examples

### Example 1
```ps
Get-NspLicenseReport.ps1 -ReportRecipient alice@example.com -ReportSender nospamproxy@example.com -ReportSubject "Example Report" -SmtpHost mail.example.com
```

### Example 2
```ps
Get-NspLicenseReport.ps1 -ReportRecipient alice@example.com -ReportSender nospamproxy@example.com -ReportSubject "Example Report" -SmtpHost mail.example.com -ReportRecipientCSV "C:\Users\example\Documents\email-recipient.csv"
```
The CSV have to contain the header "Email" else the mail addresses cannot be read from the file.  

E.g: email-recipient.csv  
User,Email  
user1,user1@example.com  
user2,user2@example.com  

The "User" header is not necessary.  

### Example 3
```ps
Get-NspLicenseReport.ps1 -NoMail -ReportType both
```
Generates a user-based and a domain-based report which are saved at the current location of execution, here: ".\"

### Example 4
```ps
Get-NspLicenseReport.ps1 -NoMail -SqlServer sql.example.com -SqlInstance NSPIntranetRole -SqlDatabase NSPIntranet
```
This generates a user-based report.  
Therefore the script connects to the SQL Server "sql.example.com" and accesses the SQL instance "NSPIntranetRole" which contains the "NSPIntranet" database.

### Example 5
```ps
Get-NspLicenseReport.ps1 -NoMail -SqlInstance ""
```
Use the above instance name "" if you try to access the default SQL instance.  
If there is a connection problem and the NSP configuration shows an empty instance for the intranet-role under "Configuration -> NoSpamProxy components" than this instance example should work.
