#
# PowerFarming.Backup
#

<#

	PowerFarming have a number of backup strategies and systems. This module is a general
    purpose toolbox of backup modules.

#>

#Symantec Backup Exec System Recovery 8.5 ( leagacy )
function Get-Besr85BackupResult([String] $driveletter, [Int32] $daysback)
{
	$evt = get-eventlog application -After (get-date).AddDays(-$daysback) | where-object {$_.Source -like 'Backup Exec System Recovery' `
			-and $_.Message -like "*$driveletter*" -and $_.Message -notlike '*6C8F1F65*'} | Select -First 1;
	return $evt;
}
export-modulemember -Function Get-Besr85BackupResult

#ShadowProtect | Unspecified Versions
function Get-ShadowProtectBackupResult([String] $volume, [Int32] $daysback)
{
	$evt = get-eventlog application -After (get-date).AddDays(-$daysback) | where-object {$_.Source -eq 'ShadowProtectSvc' `
			-and $_.Message -like "*$volume*"} | select -first 1;
	return $evt;
}
export-modulemember -Function Get-ShadowProtectBackupResult

function Get-BackupRetentionFolders([String] $backuppath)
{
	$fl = New-Object System.Collections.ArrayList
	$fldrs = Get-ChildItem -Path $backuppath -Filter BackupRetentionJob.xml -Recurse -ErrorAction SilentlyContinue -Force
	foreach ($fldr in $fldrs)
	{
		$fl.Add($fldr.Directory)
	}
	return $fl
}
export-modulemember -Function Get-BackupRetentionFolders

function Show-BackupRetentionConfigSettings([String] $folder)
{
	$flr = Get-Item -Path $folder
	[xml]$XmlDocument = Get-Content -Path "$($flr.FullName)\BackupRetentionJob.xml"
	$AgeType = $xmldocument | Select-Xml -XPath "/BackupRetention/AgeType" | Select-Object -ExpandProperty Node
	$AgeDays = $xmldocument | Select-Xml -XPath "/BackupRetention/AgeDays" | Select-Object -ExpandProperty Node
	$DeleteFileMaskInclude = $xmldocument | Select-Xml -XPath "/BackupRetention/DeleteFileMaskInclude" | Select-Object -ExpandProperty Node

	[int]$AgeDays = [convert]::ToInt32($AgeDays.InnerXml)

	Write-Host " AgeType:$($AgeType.InnerXml)"
	Write-Host " AgeDays:$($AgeDays) |"(get-date).AddDays($AgeDays* -1).ToString("dd-MM-yyyy")	
	Write-Host " DeleteFileMaskInclude:$($DeleteFileMaskInclude.InnerXml)"

}
export-modulemember -Function Show-BackupRetentionConfigSettings

function Get-BackupRetentionFileList([String] $folder)
{
	Write-Host ">>"$folder
	$flr = Get-Item -Path $folder
		[xml]$XmlDocument = Get-Content -Path "$($flr.FullName)\BackupRetentionJob.xml"
	$AgeType = $xmldocument | Select-Xml -XPath "/BackupRetention/AgeType" | Select-Object -ExpandProperty Node
	$AgeDays = $xmldocument | Select-Xml -XPath "/BackupRetention/AgeDays" | Select-Object -ExpandProperty Node
	$DeleteFileMaskInclude = $xmldocument | Select-Xml -XPath "/BackupRetention/DeleteFileMaskInclude" | Select-Object -ExpandProperty Node

	$AgeType = $AgeType.InnerXml
	[int]$AgeDays = [convert]::ToInt32($AgeDays.InnerXml)
	$oldestDate = (get-date).AddDays($AgeDays* -1)
	$DeleteFileMaskInclude = $DeleteFileMaskInclude.InnerXml
	$files = Get-ChildItem -Path $folder -Filter $DeleteFileMaskInclude | Where-Object{!($_.PSIsContainer) -and $_.CreationTime -le $oldestDate}	
	return $files
}
export-modulemember -Function Get-BackupRetentionFileList

function Export-WhatIfBackupRetentionFileList([String] $folder)
{
	$files = Get-BackupRetentionFileList([String] $folder)
	if ($files)
	{		
		$xmlWr = New-Object System.XMl.XmlTextWriter("$folder\WhatIfBackupRetentionFileList-$(Get-Date -Format 'yyyy-MM-dd-hhmmss').xml",$null)	
		$xmlWr.Formatting = 'Indented'
		$xmlWr.Indentation = 1
		$xmlWr.IndentChar = "`t"
		$xmlWr.WriteStartDocument()
		$xmlWr.WriteStartElement('BackupRetentionWhatIfFileList')
		$xmlWr.WriteStartElement('DeleteFiles')
		foreach ($file in $files)
		{
			$xmlWr.WriteStartElement('DeleteFile')
			$xmlWr.WriteAttributeString('Name', $file.Name)
			$span = New-TimeSpan (get-date) $file.CreationTime
			$xmlWr.WriteAttributeString('AgeDays', $span.Days)
			$xmlWr.WriteAttributeString('FileDate', $file.CreationTime.ToString("yyyyMMdd"))
			$xmlWr.WriteEndElement()
		}
		$xmlWr.WriteEndElement()
		$xmlWr.WriteEndDocument()
		$xmlWr.Flush()
		$xmlWr.Close()
	}
}
export-modulemember -Function Export-WhatIfBackupRetentionFileList

function Export-WhatIfBackupRetentionFileListRecurse([String] $folder)
{
	$fl = Get-BackupRetentionFolders([String] $folder)
	foreach ($fldr in $fl)
	{
		if(Test-Path -path $($fldr.ToString()))
		{
			Export-WhatIfBackupRetentionFileList([String] $fldr.ToString())	
		}
	}
}
export-modulemember -Function Export-WhatIfBackupRetentionFileListRecurse
