param (
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$true)][string] $ReportSender,
	[Parameter(Mandatory=$true)][string] $ReportRecipient,
	[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung",
	[Parameter(Mandatory=$false)][string] $SqlServer = "(local)\NoSpamProxy",
	[Parameter(Mandatory=$false)][pscredential] $Credential,
	[Parameter(Mandatory=$false)][string] $Database = "NoSpamProxyAddressSynchronization",
	[Parameter(Mandatory=$false)][bool] $TreatUnkownAsSpam = $true,
    [Parameter(Mandatory=$false)][int] $TopAddressesCount = 10,
    [Parameter(Mandatory=$false)][string[]] $ExcludeFromTopAddresses = @()
)
$reportFileName = [System.IO.Path]::Combine($Env:TEMP, "reject-analysis.html")
$totalRejected = 0
$tempRejected = 0
$permanentRejected = 0
$rblRejected = 0
$cyrenSpamRejected = 0
$cyrenAVRejected = 0
$surblRejected = 0
$characterSetRejected = 0
$wordRejected = 0
$contentrejected = 0
$greylistrejected = 0
$rdnsPermanentRejected = 0
$decryptPolicyRejected = 0
$dateStart = (Get-Date).AddDays(-$NumberOfDaysToReport)
$dateTo = Get-Date -format "dd.MM.yyyy HH:mm:ss"
$dateFrom = $dateStart.ToString("dd.MM.yyyy HH:mm:ss")

function sumUp($tmpArray){
	[int]$tmpValue = 0
	foreach($value in $tmpArray){
		$tmpValue = $tmpValue + [int]$value
	}
	return $tmpValue
}

function New-DatabaseConnection() {
	$connectionString = "Server=$SqlServer;Database=$Database;"
	if ($Credential) {
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

function Coalesce-Zero($a) {
	if ($a) { $a } else { 0 } 
}

function Invoke-SqlQuery([string] $queryName, [bool] $isInlineQuery = $false, [bool] $isSingleResult) {
	try {
		$connection = New-DatabaseConnection
		$command = $connection.CreateCommand()
		if ($isInlineQuery) {
			$command.CommandText = $queryName;
		}
		else {
			$command.CommandText = (Get-Content "$queryName.sql") -f $dateFrom, $dateTo	
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

"Getting List of mails to unknown recipients..."
$databaseVersion = [Version] (Invoke-SqlQuery "SELECT value FROM sys.fn_listextendedproperty ('AddressSynchronizationDBVersion', null, null, null, null, null, default)" -isInlineQuery $true -isSingleResult $true)

if ($databaseVersion -gt ([Version] "11.2.22")) {
	$MailsToInvalidRecipients = Invoke-SqlQuery "UnknownRecipients_Current" -isSingleResult $true
}
else {
	$MailsToInvalidRecipients = Invoke-SqlQuery "UnknownRecipients_Old" -isSingleResult $true
}


"Getting MessageTracks..."
$blockedMessageStatistics = Invoke-SqlQuery "BlockedMessageTracks"
"Getting actions statistics..."
$actions = Invoke-SqlQuery "Actions"
"Getting filter statistics..."
$filters = Invoke-SqlQuery "Filters"
"Getting Address Counts"
$addresses = Invoke-SqlQuery "TopAddresses"
$spammers = Invoke-SqlQuery "TopSpammers"
"Building report."

$totalMails = $blockedMessageStatistics | Where-Object {$_.Direction -eq "Summary" -and $_.Status -eq "Summary"} | Select-Object -ExpandProperty Count -First 1
$tempRejected = $blockedMessageStatistics | Where-Object {$_.Direction -eq "Inbound" -and ($_.Status -eq "Temporary Blocked")} | Select-Object -ExpandProperty Count -First 1
$permanentRejected = $blockedMessageStatistics | Where-Object {$_.Direction -eq "Inbound" -and ($_.Status -eq "Permanently Blocked")} | Select-Object -ExpandProperty Count -First 1
if ($null -eq $tempRejected){
	$tempRejected = 0
}
if ($null -eq $permanentRejected){
	$permanentRejected = 0
}
$totalRejected = $tempRejected + $permanentRejected
$totalRejected += $blockedMessageStatistics | Where-Object {$_.Direction -eq "Inbound" -and ($_.Status -eq "Temporary Blocked")} | Select-Object -ExpandProperty Count -First 1
$inboundmessages = $blockedMessageStatistics | Where-Object {$_.Direction -eq "Inbound" -and $_.Status -eq "Summary"} | Select-Object -ExpandProperty Count -First 1
$outboundmessages = $blockedMessageStatistics | Where-Object {$_.Direction -eq "Outbound" -and $_.Status -eq "Summary"} | Select-Object -ExpandProperty Count -First 1

$rblRejected = Coalesce-Zero ($filters |  Where-Object {$_.Name -eq "realtimeBlocklist" } | Select-Object -ExpandProperty Count -First 1)
$surblRejected = Coalesce-Zero ($filters |  Where-Object {$_.Name -eq "surblFilter" } | Select-Object -ExpandProperty Count -First 1)
$cyrenSpamRejected = Coalesce-Zero ($filters |  Where-Object {$_.Name -eq "cyrenFilter" } | Select-Object -ExpandProperty Count -First 1)
$cyrenIPRejected = Coalesce-Zero ($filters |  Where-Object {$_.Name -eq "cyrenIpReputationFilter" } | Select-Object -ExpandProperty Count -First 1)
$characterSetRejected = Coalesce-Zero ($filters |  Where-Object {$_.Name -eq "characterSetFilter" } | Select-Object -ExpandProperty Count -First 1)
$wordRejected = Coalesce-Zero ($filters |  Where-Object {$_.Name -eq "wordFilter" } | Select-Object -ExpandProperty Count -First 1)
$rdnsPermanentRejected = Coalesce-Zero ($filters |  Where-Object {$_.Name -eq "reputation" } | Select-Object -ExpandProperty Count -First 1)
$cyrenAVRejected = Coalesce-Zero (sumUp ($actions |  Where-Object {$_.Name -eq "cyrenAction" } | Select-Object -ExpandProperty Count))
$contentrejected = Coalesce-Zero (sumUp ($actions |  Where-Object {$_.Name -eq "ContentFiltering" } | Select-Object -ExpandProperty Count))
$decryptPolicyRejected = Coalesce-Zero ($actions |  Where-Object {$_.Name -eq "validateSignatureAndDecrypt" } | Select-Object -ExpandProperty Count -First 1)

$topSpammers = $spammers | select -first $TopAddressesCount

$ownDomains = (Get-NspOwnedDomain).Domain
$topRecipientsOutgoing = ($addresses | ?{$_.AddressType -eq "Recipient" -and $_.Address -notin $ExcludeFromTopAddresses -and $_.Domain -in $ownDomains} | Sort Count -Descending | select -First $TopAddressesCount)
$topSendersOutgoing = ($addresses | ?{$_.AddressType -eq "Sender" -and $_.Address -notin $ExcludeFromTopAddresses -and $_.Domain -in $ownDomains} | Sort Count -Descending | select -First $TopAddressesCount)
$topSendersIncoming = ($addresses | ?{$_.AddressType -eq "Recipient" -and $_.Address -notin $ExcludeFromTopAddresses -and $_.Domain -notin $ownDomains} | Sort Count -Descending | select -First $TopAddressesCount)
$topRecipientsIncoming = ($addresses | ?{$_.AddressType -eq "Sender" -and $_.Address -notin $ExcludeFromTopAddresses -and $_.Domain -notin $ownDomains} | Sort Count -Descending | select -First $TopAddressesCount)

"Retrieving number of mails with invalid recipients"
if ($TreatUnkownAsSpam) {
	$SpamRejected = $totalRejected - $MailsToInvalidRecipients
}
else{
	$SpamRejected = $totalRejected
}

$mailsprocessed = $totalMails

if ($inboundmessages -eq 0 -OR $null -eq $inboundmessages) {
    $blockedpercentage = 0
    $MailsToInvalidRecipientsPercentage = 0
} else {
    $blockedpercentage = [Math]::Round($SpamRejected/$inboundmessages*100,2)
    $MailsToInvalidRecipientsPercentage = [Math]::Round($MailsToInvalidRecipients/$inboundmessages*100,2)
}

if ($SpamRejected -eq 0) {
    $cyrenspamblockpercentage = 0
    $cyrenavblockpercentage = 0
    $cyrenIPBlockpercentage = 0
    $surblblockedpercentage = 0
    $charactersetblockedpercentage = 0
    $wordrejectedblockedpercentage = 0
    $decryptpolicyblockedpercentage = 0
    $rblRejectedpercentage = 0
    $reputationFilterRejectedpercentage = 0
    $contentrejectedpercentage = 0
    $greylistrejectedpercentage = 0
} else {
    $cyrenspamblockpercentage = [Math]::Round($cyrenSpamRejected/$SpamRejected*100,2)
    $cyrenavblockpercentage = [Math]::Round($cyrenAVRejected/$SpamRejected*100,2)
    $cyrenIPBlockpercentage = [Math]::Round($cyrenIPRejected/$SpamRejected*100,2)
    $surblblockedpercentage = [Math]::Round($surblRejected/$SpamRejected*100,2)
    $charactersetblockedpercentage = [Math]::Round($characterSetRejected/$SpamRejected*100,2)
    $wordrejectedblockedpercentage = [Math]::Round($wordRejected/$SpamRejected*100,2)
    $decryptpolicyblockedpercentage = [Math]::Round($decryptPolicyRejected/$SpamRejected*100,2)
    $rblRejectedpercentage = [Math]::Round($rblRejected/$SpamRejected*100,2)
    $reputationFilterRejectedpercentage = [Math]::Round($rdnsPermanentRejected/$SpamRejected*100,2)
    $contentrejectedpercentage = [Math]::Round($contentRejected/$SpamRejected*100,2)
    $greylistrejectedpercentage = [Math]::Round($greylistRejected/$SpamRejected*100,2)
}


Write-Host " "
Write-Host "TemporaryReject Total:" $tempRejected
Write-Host "PermanentReject Total:" $permanentRejected
Write-Host "TotalReject:" $totalRejected
Write-Host "Unknown recipients": $MailsToInvalidRecipients
Write-Host " "
Write-Host "Sending E-Mail to " $ReportRecipient "..."

$global:htmlout = "<html>
		<head>
			<title>Auswertung der abgewiesenen E-Mails</title>
			<style>







      			table, td, th { border: 1px solid #00cc00; border-collapse: collapse; }
				th, td {padding-left:1em; padding-right:1em;}
				td:not(:first-child){text-align:right;}
				th {color:white;}
				#headerzeile         {background-color: #00cc00;}
    		</style>
		</head>
	<body style=font-family:arial>
		<table>
			<tr id=headerzeile><th>"+ $dateFrom +" bis "+ $dateTo +" ("+$NumberOfDaysToReport+" Tage)</th><th>Count</th><th>Percent</th></tr>
			<tr><td>Mails Processed</td><td>" + $mailsprocessed +"</td><td>&nbsp;</td></tr>
			<tr><td>Sent</td><td>" + $outboundmessages +"</td><td>&nbsp;</td></tr>
			<tr><td>Received</td><td>" + $inboundmessages +"</td><td>&nbsp;</td></tr>
			<tr><td>Mails to invalid recipients</td><td>" + $MailsToInvalidRecipients +"</td><td>" + $MailsToInvalidRecipientsPercentage + " %</td></tr>
			<tr><td>Mails blocked due to Spam, Virus or Policy violation</td><td>" + $SpamRejected +"</td><td>" + $blockedpercentage +" %</td></tr>
			<tr><td>Realtime Blocklist Check</td><td>" + $rblRejected +"</td><td>" + $rblRejectedpercentage +" %</td></tr>
			<tr><td>Reputation Check</td><td>" + $rdnsPermanentRejected +"</td><td>" + $reputationFilterRejectedpercentage +" %</td></tr>
			<tr><td>Cyren IP Reputation</td><td>" + $cyrenIPRejected +"</td><td>" + $cyrenIPBlockpercentage +" %</td></tr>
			<tr><td>Cyren AntiSpam</td><td>" + $cyrenSpamRejected +"</td><td>" + $cyrenspamblockpercentage +" %</td></tr>
			<tr><td>Cyren AntiVirus</td><td>" + $cyrenAVRejected +"</td><td>" + $cyrenavblockpercentage +" %</td></tr>
			<tr><td>Spam URI Realtime Blocklists</td><td>" + $surblRejected +"</td><td>" + $surblblockedpercentage +" %</td></tr>
			<tr><td>Allowed Unicode Character Sets</td><td>" + $characterSetRejected +"</td><td>" + $charactersetblockedpercentage +" %</td></tr>
			<tr><td>Word Matching</td><td>" + $wordRejected +"</td><td>" + $wordrejectedblockedpercentage +" %</td></tr>
			<tr><td>DecryptPolicy Reject</td><td>" + $decryptPolicyRejected +"</td><td>" + $decryptpolicyblockedpercentage +" %</td></tr>
			<tr><td>ContentFiltering</td><td>" + $contentrejected + "</td><td>" + $contentrejectedpercentage + " %</td></tr>
			<tr><td>Greylisting</td><td>" + $greylistrejected + "</td><td>" + $greylistrejectedpercentage + " %</td></tr>
		</table>
	</body>
	</html>"




$htmlout | Out-File $reportFileName
"Sending report to $ReportRecipient"
Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $ReportRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der Auswertung der abgewiesenen E-Mails." -Attachments $reportFileName
Write-Host "Doing some cleanup.."
Remove-Item $reportFileName
Write-Host "Done."

# SIG # Begin signature block
# MIIbigYJKoZIhvcNAQcCoIIbezCCG3cCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU8syVo0Qz9N5qbvfyN4jaj25D
# 9qygghbWMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
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
# AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUo04q1ZTl04hT873D
# Thf89WyDNQswDQYJKoZIhvcNAQEBBQAEggEAr4plTeGv0ahAsG6zN2BQDnPsQKK5
# LYDwjdjrUvdwY9joXWeqX1xLapr45UtHsO1e15Ci8aSw8Gd/fv9MAPTv0SNxicm2
# R3t+NmkUhFijxHgElUqBSmyJBODAG+EcFOIIC/zvB9ATuhw18Keook1fIBXKn3dQ
# 6XCBZHDHV3Erg1qpsbB/hy0+YBe/8GXMTSrLIRDOwmkhFoBMNCfVgX0xeSFLhxd+
# O0DQWN+Q/LZ+OinGirDiDWKPgmobOKKe1kT0KKxPpNsfRiJro0a+o6q3Yq9h+Azo
# 8NdQLOXzJ5u29aUC1NkdI3Yfy8OzREsKISmim4qMVWmXYNuWTd+wIChLUKGCAg8w
# ggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfw
# ZjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTkwNzI0MTI1NTI2WjAjBgkqhkiG9w0BCQQxFgQUgDnX/ugbdxBm
# gFRKpVzWNm0gUMswDQYJKoZIhvcNAQEBBQAEggEAUWniERwl/+XiXThmjrDhPmHn
# fkS4E9Mpx0U4nIW0ImeDQtfjMa9EZzA9x1RqBtsGeW7Gr65HmR0g1N9aTYEXvKtd
# TTm/cvppGgZjlT2RulD5/K/igpjXee77KnfCHXPj8bttI3qucPn0nsnx8Ln0Wm4N
# mqWYurkGrle0UJkBgrstOa/NHZxAMtZgrd+zYaQV/+CJLyxr+WpbkTF3EhVQ3Hh2
# 903JHkzBy9bE6vtxWUzEkx3WrR6QibPnPYYuwlSFcuD9dJik/DCotKtFKhMPysOO
# TQ1aY3rF7wEVyotgiyQVyG0wrgYQazKqkTCCQV24qqol20AhGICoWFD8IFbx+w==
# SIG # End signature block
