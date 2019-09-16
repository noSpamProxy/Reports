param (
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$true)][string] $ReportSender,
	[Parameter(Mandatory=$true)][string] $ReportRecipient,
	[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung",
	[Parameter(Mandatory=$false)][int] $TopAddressesCount = 5,
	[Parameter(Mandatory=$false)][string[]] $excludeFromTopAddresses = @()
)

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
# SIG # Begin signature block
# MIIbigYJKoZIhvcNAQcCoIIbezCCG3cCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUckLOLFuvUawr/V9Pj0mc/VXc
# aFGgghbWMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
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
# AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU6R/5wm+euUVScZPz
# bKja0YkaqAYwDQYJKoZIhvcNAQEBBQAEggEAlT3PkeL/42rGH04xAzbWB3b3VF00
# 0HdM62qWyqrHg8k/OLH47Ey2tcbgW256dy2eHebnwd4x1SWWZjxxxKTdn04nP6No
# 9F7SBPR3x9otZp2N1w/VZcHTWMc0daZCdItZvF/lUt5TR+Sb2liQNyRe7WMxo6nV
# dco0DqEdbw466pO6ac8lvXSc+Ks7TPttngrGJhgFnFbcAkHipV/aILt7dLo0/Mse
# qsivNVEWRrMbOz5ZjjLblIODIp3DLwlstMlihak3oo0IlaiFkLgjBhugWhl830FC
# SbfbmoKTT5prA86lZvQV68SWdXed3QJRGN7YiRYj5lEzWI1j1+JmPIgtZqGCAg8w
# ggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfw
# ZjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTkwOTE2MTM0NjQwWjAjBgkqhkiG9w0BCQQxFgQU7ID/71yoeTlU
# BAS86rOANpT4BlcwDQYJKoZIhvcNAQEBBQAEggEAFwvNP4V9tWM6McuDKbGxPL/B
# d2DW1zoshMGvRQM+7su5QrGf9j7zrgZDFOA5vZSIfTY0uOCN6Z9silCOq21cZXo0
# 3f6GhUebdh6XhbJ73/tzMDP2ePnezqNu5JJ3i/m3ZPaXkPyfCSO5olsj8/cx0orF
# 3br2cqiKvMARNhod96qG2oZqDo5PRGWAv0eB4S5T7Wq/BLjDKAQhfl5ctVnPDqie
# 7+9rK2El3gZxBDWUhaOITjDSMRiwpMJuKwpTHYMi942mIkpb5BEQtXi9BST4xROA
# Ey+ZDyHPP2FbaaT6T4YCJBVn4RITAwWig8sB7w1e/HNqfZfClkm/59nn6XcXeg==
# SIG # End signature block
