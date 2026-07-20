
# Get-NspMsgTrackAsCSV.ps1

Exports filtered message tracks to a CSV containing the following information:

 - sent time
 - subject
 - status
 - sender
 - recipient
 - signed
 - encrypted
 - message ID
 

## Authentication
The script signs in with `Connect-Nsp` and supports BASIC or integrated Windows (**NTLM**) authentication.

- **Credential**: Optional `PSCredential`. When set, BASIC authentication is used.
- **CredentialFileLocation**: Optional path to a saved BASIC credential file. The default is `NspBasicAuthCredential.bin` beside the script. Relative paths such as `./auth.bin` are resolved beside the script, not from the current PowerShell directory.
- **SetCredential**: Creates or replaces the credential file and exits. It is a standalone parameter set, so no report parameters are required.
- **UseIntegratedAuth**: Forces NTLM and ignores both `Credential` and the saved credential file.

Credential selection order is: `UseIntegratedAuth` (NTLM), explicit `Credential` (BASIC), saved credential file (BASIC), then NTLM when no credential is available. The saved file contains the username and a password encrypted with Windows DPAPI `LocalMachine`; it can be read by users on the same computer who have filesystem permission to read the file.

Create the default credential file:
```ps
.\Get-NspMsgTrackAsCSV.ps1 -SetCredential
```

Create a credential file at a custom script-relative path and use it:
```ps
.\Get-NspMsgTrackAsCSV.ps1 -SetCredential -CredentialFileLocation .\auth.bin
.\Get-NspMsgTrackAsCSV.ps1 -NoTime -CredentialFileLocation .\auth.bin
```

Hardcoded credential (for automation):
```ps
$pass = ConvertTo-SecureString 'PASSWORD' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('DOMAIN\\user', $pass)
.\Get-NspMsgTrackAsCSV.ps1 -Credential $cred
```

## Usage 

```ps
Get-NspMsgTrackAsCSV.ps1 -FromDate <String> [-ToDate <String>] [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [-TenantPrimaryDomain <String>] [-TransactionId <String>] [-IpAddress <String>] [-MailSubject <String>] [-MessageId <String>] [-MinScl <Int32>] [-MaxScl <Int32>] [-Direction <String>] [-EmailAddress1 <String>] [-EmailAddress2 <String>] [<CommonParameters>]
```
```ps
Get-NspMsgTrackAsCSV.ps1 -NumberOfDaysToReport <String> [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [-TenantPrimaryDomain <String>] [-TransactionId <String>] [-IpAddress <String>] [-MailSubject <String>] [-MessageId <String>] [-MinScl <Int32>] [-MaxScl <Int32>] [-Direction <String>] [-EmailAddress1 <String>] [-EmailAddress2 <String>] [<CommonParameters>]
```
```ps
Get-NspMsgTrackAsCSV.ps1 -NumberOfHoursToReport <Int32> [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [-TenantPrimaryDomain <String>] [-TransactionId <String>] [-IpAddress <String>] [-MailSubject <String>] [-MessageId <String>] [-MinScl <Int32>] [-MaxScl <Int32>] [-Direction <String>] [-EmailAddress1 <String>] [-EmailAddress2 <String>] [<CommonParameters>]
```
```ps
Get-NspMsgTrackAsCSV.ps1 -ReportInterval <String> [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [-TenantPrimaryDomain <String>] [-TransactionId <String>] [-IpAddress <String>] [-MailSubject <String>] [-MessageId <String>] [-MinScl <Int32>] [-MaxScl <Int32>] [-Direction <String>] [-EmailAddress1 <String>] [-EmailAddress2 <String>] [<CommonParameters>]
```
```ps
Get-NspMsgTrackAsCSV.ps1 -NoTime [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [-TenantPrimaryDomain <String>] [-TransactionId <String>] [-IpAddress <String>] [-MailSubject <String>] [-MessageId <String>] [-MinScl <Int32>] [-MaxScl <Int32>] [-Direction <String>] [-EmailAddress1 <String>] [-EmailAddress2 <String>] [<CommonParameters>]
```

## Parameters
### Credential
Optional `PSCredential` for NoSpamProxy BASIC authentication. It takes precedence over a saved credential file.

### CredentialFileLocation
Optional path to the saved NoSpamProxy BASIC credential file. The default is `NspBasicAuthCredential.bin` in the script directory; relative paths are also resolved from the script directory.

### SetCredential
Creates or replaces the saved NoSpamProxy credential file and then exits. Prompts for the NoSpamProxy username and password; no report parameters are required.

### UseIntegratedAuth
Forces integrated Windows (NTLM) authentication and ignores `Credential` and `CredentialFileLocation`.

### Direction
  Default: All
  Filter for the direction of an email. (Inbound/Outbound)
  Possible values:
  All | FromExternal | FromLocal | None

### EmailAddress1
  The email address / part of the address of one of the communication participants.

### EmailAddress2
  The email address / part of the address of one of the communication participants.

### FromDate
  Mandatory if you like to use a timespan.  
  Specifies the start date for the E-Mail filter.  
  Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"  
  E.g.:  
  	"2019-06-05 08:00" or "2019-06-05 20:00:00"  

### NoMail
Does not generate an E-Mail and saves the generated reports to the current execution location.  
Ideal for testing or manual script usage.  

### IpAddress
  Filter for an IP address.

### MailSubject
  Filter for the subject / part of a subject of an email.

### MaxScl
  Filter for messages which have a maximal SCL value.

### MessageId
  Fitler for the message id of an email.

### MinScl
  Filter for messages which have a minimal SCL value.

### NoTime
  Mandatory if you do not like to specify a time value in any kind of way.  
  No value needs to be passed here \<NoTime> is just a single switch.  
  
### NspRule
  Specify a rule name which is defined in NSP as E-Mail filter.

### NumberOfDays
  Mandatory if you like to use a number of days for filtering.  
  Specifies the number of days for which th E-Mails should be filtered.  

### NumberOfHoursToReport
  Mandatory if you like to use a number of hours for filtering.  
  Specifies the number of hours for which th E-Mails should be filtered.  

### ReportFileName
**Default:** msgTrackReport  
Define a part of the complete file name.  
 
E.g.: 2019-06-05-msgTrackReport.csv  
	
### ReportInterval
Mandatory if you like to use a predefined timespan.  
Specifies a predefined timespan.  

Possible values are:  
daily, monthly, weekly  

The report will start at 00:00:00 o'clock and ends at 23:59:59 o'clock.  
The script call must be a day after the desired report end day.  

### ReportRecipient
Specifies the E-Mail recipient. It is possible to pass a comma separated list to address multiple recipients.  
E.g.: alice@example.com, bob@example.com

### ReportRecipientCSV
Set a filepath to an CSV file containing a list of report E-Mail recipient. Be aware about the needed CSV format, please watch the provided example.

### ReportSender
**Default:** NoSpamProxy Report Sender \<nospamproxy@example.com>  
Sets the report E-Mail sender address.
  
### ReportSubject
**Default:** Message Track Report   
Sets the report E-Mail subject.
	
### SmtpHost
Specifies the SMTP host which should be used to send the report E-Mail.  
It is possible to use a FQDN or IP address.
	
### Status
Specifies a filter to get only E-Mails which are matching the defined state.  

Possible values are:  
None | Success | DispatcherError | TemporarilyBlocked | PermanentlyBlocked | PartialSuccess | DeliveryPending | Suppressed | DuplicateDrop | PutOnHold | All

### TenantPrimaryDomain
Used to login into the desired NoSpamProxy tenant to run this script on.  
Only required if NoSpamProxy v14 is used in provider mode.
  
### ToDate
Optional if you like to use a timespan.  
Specifies the end date for the E-Mail filter.  
Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"  
E.g.:  
  "2019-06-05 08:00" or "2019-06-05 20:00:00"
	
### TransactionId
  Filter for the transaction ID which is provided by a NoSpamProxy Messagetrack.

## Outputs
Report is stored under %TEMP% if the report is send via E-Mail the file will be removed.


## Examples
### Example 1
```ps
Get-NspMsgTrackAsCSV.ps1 -NoTime -Status "Success" -ReportFileName "Example-Report" -ReportRecipient alice@example.com -ReportSender "NoSpamProxy Report Sender \<nospamproxy@example.com>" -ReportSubject "Example Report" -SmtpHost mail.example.com
```

### EXAMPLE 2
```ps
Get-NspMsgTrackAsCSV.ps1 -FromDate: "2019-06-05 08:00:00" -ToDate: "2019-06-05 20:00:00"  -NoMail
```
It is mandatory to specify \<FromDate>. Instead \<ToDate> is optional.  
These parameters can be combined with all other parameters except \<NumberOfDaysToReport>, \<NumberOfHoursToRepor>, \<ReportIntervall> and \<NoTime>.

### EXAMPLE 3
```ps
Get-NspMsgTrackAsCSV.ps1 -NumberOfDaysToReport 7 -NoMail
```
You can combine \<NumberOfDaysToReport> with all other parameters except \<FromDate>, \<ToDate>, \<NumberOfHoursToRepor>, \<ReportIntervall> and \<NoTime>.
  
### EXAMPLE 4
```ps
Get-NspMsgTrackAsCSV.ps1 -NumberOfHoursToReport 12 -NoMail
```
You can combine \<NumberOfHoursToReport> with all other parameters except \<FromDate>, \<ToDate>, \<NumberOfDaysToReport>, \<ReportIntervall> and \<NoTime>.
	
### EXAMPLE 5
```ps
Get-NspMsgTrackAsCSV.ps1 -ReportInterval weekly -NoMail
```
You can combine \<ReportInterval> with all other parameters except \<FromDate>, \<ToDate>, \<NumberOfDaysToReport>, \<NumberOfHoursToReport>, \<ReportIntervall> and \<NoTime>.
  
### EXAMPLE 6
```ps
Get-NspMsgTrackAsCSV.ps1 -NoTime -NoMail -NspRule "All other inbound mails"
```

### EXAMPLE 7
```ps
Get-NspMsgTrackAsCSV.ps1 -NoTime -SmtpHost mail.example.com -ReportRecipientCSV "C:\Users\example\Documents\email-report.csv"
```
The CSV have to contain the header "Email" else the mail addresses cannot be read from the file.  
It is possible to combine \<ReportRecipientCSV> with \<ReportRecipient>.  
E.g: email-report.csv  
User,Email  
user1,user1@example.com  
user2,user2@example.com  