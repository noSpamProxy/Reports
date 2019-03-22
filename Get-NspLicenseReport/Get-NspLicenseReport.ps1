param (
	[Parameter(Mandatory=$true)][string] $ExportFilePath
)
$minver = [Version]4.0.0;
if ($PSVersionTable.PSVersion -gt $minver)
{
$features = Get-nspfeatureusage;
$protcounter = 0;
$protarray = @();
$enccounter = 0;
$encarray = @();
$lfcounter = 0;
$lfarray = @();
$disccounter = 0;
$discarray = @();
$sandboxcounter = 0;
$sandboxarray = @();
foreach ($user in $features){
	if ($user.Protection -gt 0)
	{
		$protcounter+=1;
		$protarray+=$user.DisplayName;
	}
	if ($user.Encryption -gt 0)
	{
		$enccounter+=1;
		$encarray+=$user.DisplayName;
	}	
	if ($user.LargeFiles -gt 0)
	{
		$lfcounter+=1;
		$lfarray+=$user.DisplayName;
	}	
	if ($user.Disclaimer -gt 0)
	{
		$disccounter+=1;
		$discarray+=$user.DisplayName;
	}	
	if ($user.FilesUploadedToSandbox -gt 0)
	{
		$sandboxcounter+=$user.FilesUploadedToSandbox;
		$sandboxarray+=$user.DisplayName + " - " + $user.FilesUploadedToSandbox + " Uploaded File(s)";
	}	
}
"Feature-Usage:" > $ExportFilePath;
"" >>$ExportFilePath;
"Protection: " + $protcounter + " User" >> $ExportFilePath;
$protarray >> $ExportFilePath;
"" >>$ExportFilePath;
"" >>$ExportFilePath;
"" >>$ExportFilePath;
"Encryption: " + $enccounter + " User" >> $ExportFilePath;
$encarray >> $ExportFilePath;
"" >>$ExportFilePath;
"" >>$ExportFilePath;
"" >>$ExportFilePath;
"LargeFiles: " + $lfcounter + " User" >> $ExportFilePath;
$lfarray >> $ExportFilePath;
"" >>$ExportFilePath;
"" >>$ExportFilePath;
"" >>$ExportFilePath;
"Disclaimer: " + $disccounter + " User" >> $ExportFilePath;
$discarray >> $ExportFilePath;
"" >>$ExportFilePath;
"" >>$ExportFilePath;
"" >>$ExportFilePath;
"Sandbox: " + $sandboxcounter + " Uploaded Files" >> $ExportFilePath;
$sandboxarray >> $ExportFilePath;
}
else 
{
Write-Host "Eror: PowerShell version is to old. Please Update PowrShell to version 4.0.0"
}