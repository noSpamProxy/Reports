# Reports
This repository contains PowerShell Scripts to create advanced reports about your NoSpamProxy installation.

## Get-RejectionReport

This report displays information about all rejected e-mails and which filters and actions were involved. Use this version for NoSpamProxy 12.x and higher.

Currently, there are three versions of this script:

* [NoSpamProxy 12.x and later, access via Database](Get-RejectionReport%20(Database))
* [NoSpamProxy 12.x and later, access via WebServices](Get-RejectionReport%20(WebServices))
* [NoSpamProxy 11.x](11.x/Get-RejectionReport)

## [Get-RejectionReport](11.x/Get-RejectionReport)

This report displays information about all rejected e-mails and which filters were involved. Use this version for NoSpamProxy 11.x and higher.

## [Send-ReporttoUserswithAllBlockedEmails](https://github.com/noSpamProxy/Reports/tree/master/Send-ReporttoUserswithAllBlockedEmails)

Sends a report to every E-Mail address that contains all permanently blocked E-Mails in a specific period of time.

## [Send-ReporttoUserswithBlockedEmailsWithAttachment](https://github.com/noSpamProxy/Reports/tree/master/Send-ReporttoUserswithBlockedEmailsWithAttachment)

Sends a report to every E-Mail address that contains all permanently blocked E-Mails in a specific period of time that have been blocked because of an attachment.

## [Get-DMARCReport](https://github.com/noSpamProxy/Reports/tree/master/get-DMARCReport)

This report displays information about DMARC Check results for all successful inbound emails.

## [Get-expiredTLSCertificates](https://github.com/noSpamProxy/Reports/tree/master/get-expiredTLSCertificates)

This report displays all expired certificates from e-mail servers that NoSpamProxy has sent outbound e-mails to.
