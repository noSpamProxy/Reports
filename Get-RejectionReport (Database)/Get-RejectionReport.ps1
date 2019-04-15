param (
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$false)][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	[Parameter(Mandatory=$true)][string] $ReportRecipient,
	[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung",
	[Parameter(Mandatory=$false)][string] $SqlServer = "(local)\NoSpamProxyDB",
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
$cyrenAVRejected = Coalesce-Zero (($actions |  Where-Object {$_.Name -eq "cyrenAction" } | Select-Object -ExpandProperty Count))
$contentrejected = Coalesce-Zero (($actions |  Where-Object {$_.Name -eq "ContentFiltering" } | Select-Object -ExpandProperty Count))
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

if ($inboundmessages -eq 0) {
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
# MIIbnQYJKoZIhvcNAQcCoIIbjjCCG4oCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU6N1J0Ai/jL9ZBSL72mtdW+6r
# BbGgghbpMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
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
# AQkEMRYEFL8oB6UbbXTRbz87OrKRXqu2m/e8MA0GCSqGSIb3DQEBAQUABIIBAEKD
# YP8BjbpEfcPMTJXbivUUiVn+8evbCvhPulR1dQDad9Af5KILRpduvH94oiZAl97j
# hksbOg0Hc+ZiDg6cTzZsg8JbHNhmnzuXtfwt52Ku34BnieJXcpOpOpjDorHTOHHP
# cDuWlmFsyDvCt+WJA7WEqttH+WlikpXJQ3WfiP8BjLSCUFFdu/k8puSE+1ECOCxK
# B7oQHcRami8Mm+CjhjXkVhVVCLCEtOZ9EtrtIF9/V0XCu+Wg6X5ny0hVqXE5hQvv
# bEcACeEVmZ1WVMXMFXtC+VYGTW0PBBA2Htlcscwr9tqTPHVxIUQS7zs5t/rTv/r+
# eXkzBgPslGUBm28YGjahggIPMIICCwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBi
# MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3
# d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENB
# LTECEAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzEL
# BgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE5MDQxNTExMzY0MVowIwYJKoZI
# hvcNAQkEMRYEFCSBN8zocRt7f8GvXaHg0psSqKuQMA0GCSqGSIb3DQEBAQUABIIB
# AGAwGF6nZ9f8qUb/Xjgt4ZJ8e46p7oIL7Pbygzeyp/6s/EtqMOactdjJJUTB6JeT
# b/7QoIFsF1QaW/dQpGqdnIvxbQTiHRqHj25Hs9yAPavSh7zM67g2ZKo5Lok/u77s
# egSfetxZhSZKQKxRyt3WieeNwH5TDd2sjdKE52JUL4Q3HW4VX6Y4fgJwNVdcAj9E
# rU71JjXjvAZnDLxahkKgshAdcbIummSZjp9F4mSupN804miIfD4M3DuUKLeR7RSA
# AFU9Y8aS26qzu5aqi7ZFzdspc7UFiDFzPaEbV8KBAxOj2DHYcxf+nvSrPXV+l6iv
# LRCYmGrXkdzo3aa0m0Ie4bc=
# SIG # End signature block
