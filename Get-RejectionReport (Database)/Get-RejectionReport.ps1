param (
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$false)][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	[Parameter(Mandatory=$true)][string] $ReportRecipient,
	[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung",
	[Parameter(Mandatory=$false)][string] $SqlServer = "(local)\NoSpamProxyDB",
	[Parameter(Mandatory=$false)][pscredential] $Credential,
	[Parameter(Mandatory=$false)][string] $Database = "NoSpamProxyAddressSynchronization",
	[Parameter(Mandatory=$false)][bool] $TreatUnkownAsSpam = $true
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
$dateTo = Get-Date -format "dd.MM.yyyy"
$dateFrom = $dateStart.ToString("dd.MM.yyyy")

function New-DatabaseConnection() {
	$connectionString = "Server=$SqlServer;Database=$Database;"
	if ($Credential) {
		$networkCredential = $Credential.GetNetworkCredential
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
	$connection = New-DatabaseConnection
	try {
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
$cyrenAVRejected = Coalesce-Zero ($actions |  Where-Object {$_.Name -eq "cyrenAction" } | Select-Object -ExpandProperty Count -First 1)
$contentrejected = Coalesce-Zero ($actions |  Where-Object {$_.Name -eq "ContentFiltering" } | Select-Object -ExpandProperty Count -First 1)
$decryptPolicyRejected = Coalesce-Zero ($actions |  Where-Object {$_.Name -eq "validateSignatureAndDecrypt" } | Select-Object -ExpandProperty Count -First 1)

"Retrieving number of mails with invalid recipients"
if ($TreatUnkownAsSpam) {
	$SpamRejected = $totalRejected - $MailsToInvalidRecipients
}
else{
	$SpamRejected = $totalRejected
}

$mailsprocessed = $totalMails
$blockedpercentage = [Math]::Round($SpamRejected/$inboundmessages*100,2)
$MailsToInvalidRecipientsPercentage = [Math]::Round($MailsToInvalidRecipients/$inboundmessages*100,2)
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
Write-Host " "
Write-Host "TemporaryReject Total:" $tempRejected
Write-Host "PermanentReject Total:" $permanentRejected
Write-Host "TotalReject:" $totalRejected
Write-Host "Unknown recipients": $MailsToInvalidRecipients
Write-Host " "
Write-Host "Sending E-Mail to " $ReportRecipient "..."

$htmlout = "<html>
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
			<tr><td>Cyren Premium AntiVirus</td><td>" + $cyrenAVRejected +"</td><td>" + $cyrenavblockpercentage +" %</td></tr>
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
