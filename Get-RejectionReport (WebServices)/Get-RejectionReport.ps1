param (
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$true)][string] $ReportSender,
	[Parameter(Mandatory=$true)][string] $ReportRecipient,
	[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung",
	[Parameter(Mandatory=$false)][int] $TopAddressesCount = 5,
	[Parameter(Mandatory=$false)][string[]] $excludeFromTopAddresses = @(),
	# only needed for v14 with enabled provider mode
	[Parameter(Mandatory=$false)][string] $TenantPrimaryDomain
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
	if ($(Get-NspIsProviderModeEnabled) -eq $true) {
		if ($null -eq $TenantPrimaryDomain -OR $TenantPrimaryDomain -eq "") {
			Write-Host "Please provide a TenantPrimaryDomain to run this script with NoSpamProxy v14 in provider mode."
			EXIT
		} else {
			# NSP v14 has a new authentication mechanism, Connect-Nsp is required to authenticate properly
			# -IgnoreServerCertificateErrors allows the usage of self-signed certificates
			Connect-Nsp -IgnoreServerCertificateErrors -PrimaryDomain $TenantPrimaryDomain
		}
	}
}

$reportFileName = [System.IO.Path]::Combine($Env:TEMP, "reject-analysis.html")
$totalRejected = 0
$tempRejected = 0
$permanentRejected = 0
$reputationRejected = 0
$rblRejected = 0
$cyrenSpamRejected = 0
$cyrenAVRejected = 0
$cyrenIPRepRejected = 0
$surblRejected = 0
$characterSetRejected = 0
$wordRejected = 0
$contentrejected = 0
$greylistrejected = 0
$rdnsPermanentRejected = 0
$decryptPolicyRejected = 0
$onBodyRejected = 0
$onEnvelopeRejected = 0
$dateStart = (Get-Date).AddDays(-$NumberOfDaysToReport)
$dateStop = Get-Date
$dateTo = Get-Date -format "dd.MM.yyyy"
$dateFrom = $dateStart.ToString("dd.MM.yyyy")
$topSpammers = @{}

Write-Host "Getting MessageTracks from NoSpamProxy for..."
Write-Host "`tTemporarily Blocked"
$messageTracks = Get-NSPMessageTrack -From $dateStart -Status TemporarilyBlocked -Directions FromExternal -WithActions -WithDeliveryAttempts -WithFilters -WithOperations -WithAddresses

$totalRejected += $messageTracks.Count
$tempRejected += $messageTracks.Count

foreach ($item in $messageTracks)
{
    $sender = ($item.Addresses|?{[string]$_.AddressType -eq "Sender"}).Address
    if($sender -ne $null) {
        $topSpammers[$sender]++
    }
	foreach ($action in $item.Actions){
		if ($action.Name -eq "greylisting" -and $action.Decision -notcontains "Pass")
		{
		    $onEnvelopeRejected++
	 	    $greylistrejected++
		}
	}
}

Write-Host "`tPermanently Blocked"
$messageTracks = Get-NSPMessageTrack -From $dateStart -Status PermanentlyBlocked -Directions FromExternal -WithActions -WithDeliveryAttempts -WithFilters -WithOperations -WithAddresses

$totalRejected += $messageTracks.Count
$permanentRejected += $messageTracks.Count

foreach ($item in $messageTracks)
{
    $sender = ($item.Addresses|?{[string]$_.AddressType -eq "Sender"}).Address
    if($sender -ne $null) {
        $topSpammers[$sender]++
    }
	foreach ($permanentvalidationentry in $item.Filters)
	{
		if ($permanentvalidationentry.Name -eq "realtimeBlocklist" -and $permanentvalidationentry.Scl -gt 0)
		{
			$rblRejected++
			$onEnvelopeRejected++
		}

		if ($permanentvalidationentry.Name -eq "surblFilter" -and $permanentvalidationentry.Scl -gt 0)
		{
			$surblRejected++
			$onBodyRejected++
		}

		if ($permanentvalidationentry.Name -eq "cyrenFilter" -and $permanentvalidationentry.Scl -gt 0)
		{
			$cyrenSpamRejected++
			$onBodyRejected++
		}

		if ($permanentvalidationentry.Name -eq "characterSetFilter" -and $permanentvalidationentry.Scl -gt 0)
		{
			$characterSetRejected++
			$onBodyRejected++
		}

		if ($permanentvalidationentry.Name -eq "wordFilter" -and $permanentvalidationentry.Scl -gt 0)
		{
			$wordRejected++
			$onBodyRejected++
		}

		if (($permanentvalidationentry.Name -eq "validateSignatureAndDecrypt") -and ($permanentvalidationentry.Decision -notcontains "Pass" ))
		{
			$decryptPolicyRejected++
			$onBodyRejected++
		}
	    
        if (($permanentvalidationentry.Name -eq "reputation") -and ($permanentvalidationentry.Scl -gt 0 ))
		{
			$reputationRejected++
			$onEnvelopeRejected++
		}
		if (($permanentvalidationentry.Name -eq "cyrenIpReputationFilter") -and ($permanentvalidationentry.Scl -gt 0 ))
		{
			$cyrenIPRepRejected++
			$onEnvelopeRejected++
		}
		
	}

    foreach ($action in $item.Actions)
    {
    	if ($action.Name -eq "cyrenAction" -and $action.Decision -notcontains "Pass")
		{
			$cyrenAVRejected++
			$onBodyRejected++
		}
		if ($action.Name -eq "malwareScan" -and $action.Decision -notcontains "Pass")
		{
			$cyrenAVRejected++
			$onBodyRejected++
		}
		if ($action.Name -eq "ContentFiltering" -and $action.Decision -notcontains "Pass")
		{
		    $onBodyRejected++
	 	    $contentrejected++
		}
    }
}
Write-Host "Evaluating top spammers"
$topSpammers = $topSpammers.GetEnumerator() | ?{$_.Name -notin $excludeFromTopAddresses} | sort Value -Descending | select -First $TopAddressesCount


$ownedDomains = (Get-NspOwnedDomain).Domain

Write-Host "Evaluating top Senders/Recipients Outgoing"
$messageTracks = (Get-NSPMessageTrack -From $dateStart -Directions FromLocal -Status Success -WithAddresses|?{($_.Addresses|?{[string]$_.AddressType -eq "Recipient" -and $_.Domain -in $ownedDomains}) -eq $null})

$outboundmessages = $messageTracks.Count
$topLocal = @{}
$topLocal["Sender"] = @{}
$topLocal["Recipient"] = @{}

foreach ($addr in ($messageTracks.Addresses)) {
	if(($addr.AddressType -like "Sender") -OR ($addr.AddressType -like "Recipient")){
		$topLocal[[string]$addr.AddressType][$addr.Address]++
	}
}

$topRecipientsOutgoing = ($topLocal["Recipient"].GetEnumerator() | ?{$_.Name -notin $excludeFromTopAddresses} | Sort Value -Descending | select -First $TopAddressesCount)
$topSendersOutgoing = ($topLocal["Sender"].GetEnumerator() | ?{$_.Name -notin $excludeFromTopAddresses} | Sort Value -Descending | select -First $TopAddressesCount)



Write-Host "Evaluating top Senders/Recipients Incoming"
$messageTracks = Get-NSPMessageTrack -From $dateStart -Directions FromExternal -Status Success -WithAddresses

$inboundmessages = $messageTracks.Count
$topExternal = @{}
$topExternal["Sender"] = @{}
$topExternal["Recipient"] = @{}

foreach ($addr in $messageTracks.Addresses) {
	if(($addr.AddressType -like "Sender") -OR ($addr.AddressType -like "Recipient")){
		$topExternal[[string]$addr.AddressType][$addr.Address]++
	}
}

$topRecipientsIncoming = ($topExternal["Recipient"].GetEnumerator() | ?{$_.Name -notin $excludeFromTopAddresses} | Sort Value -Descending | select -First $TopAddressesCount)
$topSendersIncoming = ($topExternal["Sender"].GetEnumerator() | ?{$_.Name -notin $excludeFromTopAddresses} | Sort Value -Descending | select -First $TopAddressesCount)



$mailsprocessed = $outboundmessages+$inboundmessages

if ($inboundmessages -eq 0) {
    $blockedpercentage = 0
} else {
    $blockedpercentage = [Math]::Round($totalRejected/$inboundmessages*100,2)
}

if ($totalRejected -eq 0) {
    $cyrenspamblockpercentage = 0
    $cyrenavblockpercentage = 0
    $surblblockedpercentage = 0
    $charactersetblockedpercentage = 0
    $wordrejectedblockedpercentage = 0
    $decryptpolicyblockedpercentage = 0
    $rblRejectedpercentage = 0
    $contentrejectedpercentage = 0
    $greylistrejectedpercentage = 0
    $reputationRejectedpercentage = 0
    $cyreniprepRejectedpercentage = 0
} else {
    $cyrenspamblockpercentage = [Math]::Round($cyrenSpamRejected/$totalRejected*100,2)
    $cyrenavblockpercentage = [Math]::Round($cyrenAVRejected/$totalRejected*100,2)
    $surblblockedpercentage = [Math]::Round($surblRejected/$totalRejected*100,2)
    $charactersetblockedpercentage = [Math]::Round($characterSetRejected/$totalRejected*100,2)
    $wordrejectedblockedpercentage = [Math]::Round($wordRejected/$totalRejected*100,2)
    $decryptpolicyblockedpercentage = [Math]::Round($decryptPolicyRejected/$totalRejected*100,2)
    $rblRejectedpercentage = [Math]::Round($rblRejected/$totalRejected*100,2)
    $contentrejectedpercentage = [Math]::Round($contentRejected/$totalRejected*100,2)
    $greylistrejectedpercentage = [Math]::Round($greylistRejected/$totalRejected*100,2)
    $reputationRejectedpercentage = [Math]::Round($reputationRejected/$totalRejected*100,2)
    $cyreniprepRejectedpercentage = [Math]::Round($cyrenIPRepRejected/$totalRejected*100,2)
}

Write-Host " "
Write-Host "TemporaryReject Total:" $tempRejected
Write-Host "PermanentReject Total:" $permanentRejected
Write-Host "TotalReject:" $totalRejected
Write-Host "Generating Report..."


$global:htmlout = "<html>
		<head>
			<title>Auswertung der abgewiesenen E-Mails</title>
			<style>
                table {border-spacing: 0px; border: 1px solid black; background-color: #3867d6; float:left; margin:10px}

                th {padding: 10px; color: white;}
      			td {padding: 6px 10px; color: white;}

                tr.newsegment>td,tr.newsegment>th {border-top-color: black; border-top-width: 1px; border-top-style: solid;}

                tr.sub>td {background-color: #4b7bec;}
                tr.sub>td:first-of-type {border-left-color: #3867d6;border-left-style:solid;border-left-width:8px}
                
    		</style>
		</head>
	<body style=font-family:arial>
		<table>
			<tr><th>"+ $dateFrom +" bis "+ $dateTo +" ("+$NumberOfDaysToReport+" Tage)</th><th>Count</th><th>Percent</th></tr>
			<tr><td>Mails Processed</td><td>" + $mailsprocessed +"</td><td>&nbsp;</td></tr>
            <tr class=`"sub`"><td>Sent</td><td>" + $outboundmessages +"</td><td>&nbsp;</td></tr>
			<tr class=`"sub`"><td>Received</td><td>" + $inboundmessages +"</td><td>&nbsp;</td></tr>

			<tr class=`"newsegment`"><td>Mails blocked</td><td>" + $totalRejected +"</td><td>" + $blockedpercentage +" %</td></tr>
			<tr class=`"sub`"><td>Realtime Blocklist Check</td><td>" + $rblRejected +"</td><td>" + $rblRejectedpercentage +" %</td></tr>
			<tr class=`"sub`"><td>Reputation Check</td><td>" + $reputationRejected +"</td><td>" + $reputationRejectedpercentage +" %</td></tr>
			<tr class=`"sub`"><td>Cyren AntiSpam</td><td>" + $cyrenSpamRejected +"</td><td>" + $cyrenspamblockpercentage +" %</td></tr>
			<tr class=`"sub`"><td>Cyren Premium AntiVirus</td><td>" + $cyrenAVRejected +"</td><td>" + $cyrenavblockpercentage +" %</td></tr>
			<tr class=`"sub`"><td>Cyren IP Reputation</td><td>" + $cyrenIPRepRejected +"</td><td>" + $cyreniprepRejectedpercentage +" %</td></tr>
			<tr class=`"sub`"><td>Spam URI Realtime Blocklists</td><td>" + $surblRejected +"</td><td>" + $surblblockedpercentage +" %</td></tr>
			<tr class=`"sub`"><td>Allowed Unicode Character Sets</td><td>" + $characterSetRejected +"</td><td>" + $charactersetblockedpercentage +" %</td></tr>
			<tr class=`"sub`"><td>Word Matching</td><td>" + $wordRejected +"</td><td>" + $wordrejectedblockedpercentage +" %</td></tr>
			<tr class=`"sub`"><td>DecryptPolicy Reject</td><td>" + $decryptPolicyRejected +"</td><td>" + $decryptpolicyblockedpercentage +" %</td></tr>
			<tr class=`"sub`"><td>ContentFiltering</td><td>" + $contentrejected + "</td><td>" + $contentrejectedpercentage + " %</td></tr>
			<tr class=`"sub`"><td>Greylisting</td><td>" + $greylistrejected + "</td><td>" + $greylistrejectedpercentage + " %</td></tr>
        </table>"


function enumerateAddressList($addrlist) {
    foreach($addr in $addrlist) {
        $global:htmlout += "<tr class=`"sub`"><td>" + $addr.Key + "</td><td>" + $addr.Value + "</td><td>&nbsp;</td></tr>"
    }
}

$global:htmlout += "<table>
            <tr><th>Top Local E-Mail Addresses</th><th>Count</th><td>&nbsp;</td></tr>
            <tr><td>Most E-Mails From</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
enumerateAddressList($topSendersOutgoing)
$global:htmlout += "<tr class=`"newsegment`"><td>Most E-Mails To</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
enumerateAddressList($topRecipientsIncoming)
$global:htmlout += "</table>"



$global:htmlout += "<table>
            <tr><th>Top External E-Mail Addresses</th><th>Count</th><td>&nbsp;</td></tr>
            <tr><td>Most E-Mails From</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
enumerateAddressList($topSendersIncoming)
$global:htmlout += "<tr class=`"newsegment`"><td>Most E-Mails To</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
enumerateAddressList($topRecipientsOutgoing)
$global:htmlout += "<tr class=`"newsegment`"><td>Top Spammers</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
enumerateAddressList($topSpammers)
$global:htmlout += "</table>"

$global:htmlout | Out-File $reportFileName

Write-Host "Report Generated Successfully"

"Sending report to $ReportRecipient"
Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $ReportRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der Auswertung der abgewiesenen E-Mails." -Attachments $reportFileName
Write-Host "Doing some cleanup.."
Remove-Item $reportFileName
Write-Host "Done."