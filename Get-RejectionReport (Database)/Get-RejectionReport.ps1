param (
	[Parameter(Mandatory=$false, ParameterSetName="default")][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$true, ParameterSetName="default")][string] $SMTPHost,
	[Parameter(Mandatory=$true, ParameterSetName="default")][string] $ReportSender,
	[Parameter(Mandatory=$true, ParameterSetName="default")][string] $ReportRecipient,
	[Parameter(Mandatory=$false, ParameterSetName="default")][string] $ReportSubject = "Auswertung",
	[Parameter(Mandatory=$false, ParameterSetName="default")][string] $SqlServer = "(local)\NoSpamProxy",
	[Parameter(Mandatory=$false, ParameterSetName="default")][string] $SqlInstance = "NoSpamProxy",
	[Parameter(Mandatory=$false, ParameterSetName="default")][pscredential] $Credential,
	[Parameter(Mandatory=$false, ParameterSetName="default")][pscredential] $SqlCredential,
	[Parameter(Mandatory=$false, ParameterSetName="default")][string] $Database = "NoSpamProxyAddressSynchronization",
	[Parameter(Mandatory=$false, ParameterSetName="default")][string] $SqlDatabase = "NoSpamProxyAddressSynchronization",
	[Parameter(Mandatory=$false, ParameterSetName="default")][bool] $TreatUnkownAsSpam = $true,
    [Parameter(Mandatory=$false, ParameterSetName="default")][int] $TopAddressesCount = 10,
    [Parameter(Mandatory=$false, ParameterSetName="default")][string[]] $ExcludeFromTopAddresses = @(),
    [Parameter(Mandatory=$false, ParameterSetName="default")][int] $TenantId = 0,
	# set SqlUser password
	[Parameter(Mandatory=$true, ParameterSetName="setLoginPassword")][switch] $SetLoginPassword,
	# sql username only used if password is saved in an encrypted binary file
    [Parameter(Mandatory = $false, ParameterSetName="default")][string] $SqlUsername,
    # locationnof the encrypted binary file for saving a sql password
    [Parameter(Mandatory=$false, ParameterSetName="default")][parameter(ParameterSetName = "setLoginPassword")][string] $SqlPasswordFileLocation = "$PSScriptRoot\NspReadSqlPass.bin",
	[Parameter(Mandatory=$true, ParameterSetName="getDepreceatedParameters")][switch] $GetDepreceatedParameters
)

#--------------------Depreceated handling----------------
# validate parameters
if ($GetDepreceatedParameters) {
	Write-Host "The following parameters are depreceated and will be removed in the future. Please replace them with their new replacements."
	Write-Host ""
	Write-Host "Database: replaced by `'SqlDatabase`'"
	Write-Host "Credential: replaced by `'SqlCredential`'"
	Write-Host "SqlServer: is splitted into `'SqlServer`' and `'SqlInstance`', please do not use a combined string anymore."
	Write-Host "If you are using the default values for parameters inside the script you do not need to change anything unless you changed those values by yourself."
	EXIT
}
$validationExit = $false
if ($PSBoundParameters.ContainsKey('SqlInstance') -AND $SqlServer.Contains("\")){
	Write-Warning "`'-SqlInstance`' cannot be used if the `'SqlServer`' parameter already contains the instance name."
	Write-Warning "Please use `'-SqlInstance`' and `'-SqlServer`' together to provider the instance name and the database server through seperate parameters."
	Write-Warning "The combined usage is depreceated and will be removed in a future version."
	Write-Host
	$validationExit = $true

}
if ($PSBoundParameters.ContainsKey('Database') -AND $PSBoundParameters.ContainsKey('SqlDatabase')){
	Write-Warning "The `'Database`' parameter is depreceated and will be removed in a future version. Please use the `'SqlDatabase`' parameter only."
	Write-Host
	$validationExit = $true
}
if ($PSBoundParameters.ContainsKey('Credential') -AND $PSBoundParameters.ContainsKey('SqlCredential')){
	Write-Warning "The `'Credential`' parameter is depreceated and will be removed in a future version. Please use the `'SqlCredential`' parameter only."
	Write-Host
	$validationExit = $true
}
if ($validationExit -eq $true){
	Write-Host "Please fix the above issues and run the script again."
	Write-Host "Call `'.\$(Split-Path $PSCommandPath -Leaf) -GetDepreceatedParameters`' to get a list of the desired replacements."
	EXIT
}

#-------------------Functions----------------------
# create an encrypted binary file which contains the password for the desired SQL user
function Set-loginPass {
    # Imports Security library for encryption
    Add-Type -AssemblyName System.Security
    $sqlPass = Read-Host -Promp 'Input your user password'
    $passFileLocation = "$SqlPasswordFileLocation"
    $inBytes = [System.Text.Encoding]::Unicode.GetBytes($sqlPass)
    $protected = [System.Security.Cryptography.ProtectedData]::Protect($inBytes, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    [System.IO.File]::WriteAllBytes($passFileLocation, $protected)
}
# return the SQL user password from the encrypted binary file
# if file does not exists give a hint and ask for manual input
function Get-loginPass {
    # Imports Security library for encryption
    Add-Type -AssemblyName System.Security
    Add-Type -AssemblyName System.Text.Encoding
    $passFileLocation = "$SqlPasswordFileLocation"
    if (Test-Path $passFileLocation) {
        try {
            $protected = [System.IO.File]::ReadAllBytes($passFileLocation)
            $rawKey = [System.Security.Cryptography.ProtectedData]::Unprotect($protected, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
            return [System.Text.Encoding]::Unicode.GetString($rawKey)
        }
        catch {
            Write-Host $_.Exception | format-list -force
        }
    }
    else {
        Write-Host "No Password file found! Please run '$($PSCommandPath) -SetLoginPassword' for saving your password encrypted."
        $loginPass = Read-Host -Promp 'Input your user password'
        return $loginPass
    }
}
function sumUp($tmpArray){
	[int]$tmpValue = 0
	foreach($value in $tmpArray){
		$tmpValue = $tmpValue + [int]$value
	}
	return $tmpValue
}

function New-DatabaseConnection() {
	$connectionString = "Server=$SqlServer\$SqlInstance;Database=$SqlDatabase;"
	if ($SqlCredential) {
		$networkCredential = $SqlCredential.GetNetworkCredential()
		$connectionString += "uid=" + $networkCredential.UserName + ";pwd=" + $networkCredential.Password + ";"
    } elseif ($SqlUsername) {
        $password = (convertto-securestring -string (Get-loginPass) -asplaintext -force)
        $Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $SqlUsername, $password
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
			$command.CommandText = (Get-Content "$queryName.sql") -f $dateFrom, $dateTo, $TenantId	
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

#-------------------Variables----------------------
# migrate old variables to new one
if ($PSBoundParameters.ContainsKey('SqlInstance') -AND $PSBoundParameters.ContainsKey('SqlDatabase')) {
	Write-Host "It looks like you are using the new parameters."
} elseif ($SetLoginPassword) {} else {
	Write-Host "Translating old paramters to new parameters."
	Write-Warning "Please adjust your script invocation to match the new parameters. The old ones will be removed in a future version."
	Write-Warning "Run `'.\$(Split-Path $PSCommandPath -Leaf) -GetDepreceatedParameters`' to get a list of the desired replacements."
	Write-Host ""
	$SqlCredential = $Credential
	$Credential = $null
	if ($SqlServer.Contains("\")){
		$SqlInstance = $SqlServer.Split('\')[1]
		$SqlServer = $SqlServer.Split('\')[0]
	} else {
		if ($PSBoundParameters.ContainsKey('SqlInstance')) {
		} else {
			Write-Warning "You provided a SQL server but no SQL instance name."
			Write-Warning "For compatibility we are trying to use `'NoSpamProxy`' for the connection."
			Write-Warning "Please use the `'-SqlInstance`' parameter for your next execution."
		}
	}
	$SqlDatabase = $Database
}

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

#--------------------Main-----------------------
# create password file for login
if($SetLoginPassword){
	Set-loginPass
	EXIT
}
$databaseVersion = [Version] (Invoke-SqlQuery "SELECT value FROM sys.fn_listextendedproperty ('AddressSynchronizationDBVersion', null, null, null, null, null, default)" -isInlineQuery $true -isSingleResult $true)
if ($databaseVersion -gt ([Version] "14.0.0")) {
	if ($TenantId -eq 0) {
		Write-Host ""
		Write-Host "If you are using the provder mode and want to run this script for a specific tenant then you need to provide the desired tenant id using the parameter `'-TenantId`'."
		Write-Host ""
	}
	#$ownDomains = Invoke-SqlQuery "SELECT Name FROM [Configuration].[OwnedDomain] WHERE TenantId = $TenantId" -isInlineQuery $true
	"Getting MessageTracks..."
	$blockedMessageStatistics = Invoke-SqlQuery "BlockedMessageTracks_v14"
	"Getting actions statistics..."
	$actions = Invoke-SqlQuery "Actions_v14"
	"Getting filter statistics..."
	$filters = Invoke-SqlQuery "Filters_v14"
	#"Getting Address Counts"
	#$addresses = Invoke-SqlQuery "TopAddresses_v14"
	#$spammers = Invoke-SqlQuery "TopSpammers_v14"
} else {
	if ($databaseVersion -gt ([Version] "11.2.22")) {
		$MailsToInvalidRecipients = Invoke-SqlQuery "UnknownRecipients_Current" -isSingleResult $true
	} else {
		$MailsToInvalidRecipients = Invoke-SqlQuery "UnknownRecipients_Old" -isSingleResult $true
	}
	#$ownDomains = (Get-NspOwnedDomain).Domain
	"Getting MessageTracks..."
	$blockedMessageStatistics = Invoke-SqlQuery "BlockedMessageTracks"
	"Getting actions statistics..."
	$actions = Invoke-SqlQuery "Actions"
	"Getting filter statistics..."
	$filters = Invoke-SqlQuery "Filters"
	#"Getting Address Counts"
	#$addresses = Invoke-SqlQuery "TopAddresses"
	#$spammers = Invoke-SqlQuery "TopSpammers"
}

"Getting List of mails to unknown recipients..."

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

<#
This code could enhance the future output of this script.

$topSpammers = $spammers | Select-Object -first $TopAddressesCount
$topRecipientsOutgoing = ($addresses | ?{$_.AddressType -eq "Recipient" -and $_.Address -notin $ExcludeFromTopAddresses -and $_.Domain -in $ownDomains} | Sort Count -Descending | select -First $TopAddressesCount)
$topSendersOutgoing = ($addresses | ?{$_.AddressType -eq "Sender" -and $_.Address -notin $ExcludeFromTopAddresses -and $_.Domain -in $ownDomains} | Sort Count -Descending | select -First $TopAddressesCount)
$topSendersIncoming = ($addresses | ?{$_.AddressType -eq "Recipient" -and $_.Address -notin $ExcludeFromTopAddresses -and $_.Domain -notin $ownDomains} | Sort Count -Descending | select -First $TopAddressesCount)
$topRecipientsIncoming = ($addresses | ?{$_.AddressType -eq "Sender" -and $_.Address -notin $ExcludeFromTopAddresses -and $_.Domain -notin $ownDomains} | Sort Count -Descending | select -First $TopAddressesCount)
#>

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
			$( if ($databaseVersion -gt ([Version] "14.0.0") -AND $cyrenIPRejected -eq 0 -AND $cyrenSpamRejected -eq 0 - $cyrenAVRejected -eq 0) {
				"<tr><td>Cyren IP Reputation</td><td>" + $cyrenIPRejected +"</td><td>" + $cyrenIPBlockpercentage +" %</td></tr>"
				"<tr><td>Cyren AntiSpam</td><td>" + $cyrenSpamRejected +"</td><td>" + $cyrenspamblockpercentage +" %</td></tr>"
				"<tr><td>Cyren AntiVirus</td><td>" + $cyrenAVRejected +"</td><td>" + $cyrenavblockpercentage +" %</td></tr>"
			})
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
