# Get-NspSandboxStatistic.ps1
Create a Cyren Sandbox statistic for rejected emails. 

## Usage 

```ps
Get-NspSandboxStatistic.ps1 -FromDate <String> [-ToDate <String>] [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [<CommonParameters>]
```
```ps
Get-NspSandboxStatistic.ps1 -NumberOfDaysToReport <String> [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [<CommonParameters>]
```
```ps
Get-NspSandboxStatistic.ps1 -NumberOfHoursToReport <Int32> [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [<CommonParameters>]
```
```ps
Get-NspSandboxStatistic.ps1 -ReportInterval <String> [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [<CommonParameters>]
```
```ps
Get-NspSandboxStatistic.ps1 -NoTime [-Status <String>] [-NspRule <String>] [-MailRecipientCSV <String>] [-NoMail] [-ReportFileName <String>] [-ReportRecipient <String[]>] [-ReportRecipientCSV <String>] [-ReportSender <String>] [-ReportSubject <String>] [-SmtpHost <String>] [<CommonParameters>]
```


## Parameters
### FromDate
  Mandatory if you like to use a timespan.  
  Specifies the start date for the E-Mail filter.  
  Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"  
  E.g.:  
  	"2019-06-05 08:00:00"  

### MailRecipientCSV
Set a file path to an CSV file containing a list of original E-Mail recipient to filter for. Be aware about the needed CSV format, please watch the provided example.  	
	
### NoMail
Does not generate an E-Mail and saves the generated reports to the current execution location.  
Ideal for testing or manual script usage.  

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
**Default:** NSP Sandbox statistic     
Sets the report E-Mail subject.
	
### SmtpHost
Specifies the SMTP host which should be used to send the report E-Mail.  
It is possible to use a FQDN or IP address.
	
### ToDate
Optional if you like to use a timespan.  
Specifies the end date for the E-Mail filter.  
Please use ISO 8601 date format: "YYYY-MM-DD hh:mm:ss"  
E.g.:  
  "2019-06-05 08:00:00"
	
## Outputs
 none

## Examples
### Example 1
```ps
Get-NspSandboxStatistic.ps1 -NoTime -ReportRecipient alice@example.com -ReportSender "NoSpamProxy Report Sender <nospamproxy@example.com>" -ReportSubject "Example Report" -SmtpHost mail.example.com
```

### EXAMPLE 2
```ps
Get-NspSandboxStatistic.ps1 -FromDate: "2019-06-05 08:00:00" -ToDate: "2019-06-05 20:00:00"  -NoMail
```
It is mandatory to specify \<FromDate>. Instead \<ToDate> is optional.  
These parameters can be combined with all other parameters except \<NumberOfDaysToReport>, \<NumberOfHoursToRepor>, \<ReportIntervall> and \<NoTime>.

### EXAMPLE 3
```ps
Get-NspSandboxStatistic.ps1 -NumberOfDaysToReport 7 -NoMail
```
You can combine \<NumberOfDaysToReport> with all other parameters except \<FromDate>, \<ToDate>, \<NumberOfHoursToRepor>, \<ReportIntervall> and \<NoTime>.
  
### EXAMPLE 4
```ps
Get-NspSandboxStatistic.ps1 -NumberOfHoursToReport 12 -NoMail
```
You can combine \<NumberOfHoursToReport> with all other parameters except \<FromDate>, \<ToDate>, \<NumberOfDaysToReport>, \<ReportIntervall> and \<NoTime>.
	
### EXAMPLE 5
```ps
Get-NspSandboxStatistic.ps1 -ReportInterval weekly -NoMail
```
You can combine \<ReportInterval> with all other parameters except \<FromDate>, \<ToDate>, \<NumberOfDaysToReport>, \<NumberOfHoursToReport>, \<ReportIntervall> and \<NoTime>.
  
### EXAMPLE 6
```ps
Get-NspSandboxStatistic.ps1 -NoTime -NoMail -NspRule "All other inbound mails"
```

### EXAMPLE 7
```ps
Get-NspSandboxStatistic.ps1 -NoTime -SmtpHost mail.example.com -ReportRecipientCSV "C:\Users\example\Documents\email-report.csv"
```
The CSV have to contain the header "Email" else the mail addresses cannot be read from the file.  
It is possible to combine \<ReportRecipientCSV> with \<ReportRecipient>.  
E.g: email-report.csv  
User,Email  
user1,user1@example.com  
user2,user2@example.com  