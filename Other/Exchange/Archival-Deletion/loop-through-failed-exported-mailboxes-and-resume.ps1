###########################################################################
#
# NAME: Spread-Mailboxes.ps1
#
# AUTHOR: Jan Egil Ring
# EMAIL: jer@powershell.no
#
# COMMENT: Script to spread mailboxes alphabetically across mailboxdatabases based on the first character in the user`s displayname.
#          For more information, see the following blog-post: http://blog.powershell.no/2010/05/14/script-to-spread-exchange-mailboxes-alphabetically-across-databases
#
# You have a royalty-free right to use, modify, reproduce, and
# distribute this script file in any way you find useful, provided that
# you agree that the creator, owner above has no warranty, obligations,
# or liability for such use.
#
# VERSION HISTORY:
# 1.0 14.05.2010 - Initial release
#
###########################################################################

#Add the Exchange Server 2010 Management Shell snapin
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

Function LogStamp 
{

	$now=get-Date
	$yr=$now.Year.ToString()
	$mo=$now.Month.ToString()
	$dy=$now.Day.ToString()
	$hr=$now.Hour.ToString()
	$mi=$now.Minute.ToString()
	
	if ($mo.length -lt 2) 
		{
		$mo="0"+$mo #pad single digit months with leading zero
		}
	if ($dy.length -lt 2) 
		{
		$dy="0"+$dy #pad single digit day with leading zero
		}
	if ($hr.length -lt 2) 
		{
		$hr="0"+$hr #pad single digit hour with leading zero
		}
	if ($mi.length -lt 2) 
		{
		$mi="0"+$mi #pad single digit minute with leading zero
		}
	#write-output $yr$mo$dy$hr$mi
	return $yr + $mo + $dy + $hr + $mi
}

$server = "PFNZ-SRV-030"
$start = '01/01/2008'
$end = '01/01/2009'
$year = '2008'
$exportfolder = "\\pfnz-srv-030\archives\"
#$exportfolder = "\\pfnz-srv-028\PFW-MAILBOX-ARCHIVES\" -- recommendations are that UNC are used, but not across the network. ie. local transfers only.
$batchname = [System.Guid]::NewGuid()

#Loop through each mailbox
foreach ($mailbox in (Get-Mailbox -server $server)) {

	$displayname = $mailbox.Displayname
	$exportid = $mailbox.Displayname + " - Mail Archive - " + $year
	$name = "AJ:" + (LogStamp) + ":" + $year + ":" + $mailbox.Alias
	#New-MailboxExportRequest -Name $name -BatchName $batchname -ContentFilter "((Received -gt '$start') -and (Received -lt '$end'))" -Mailbox $displayname -FilePath ($exportfolder + $exportid)

#if ($displayname -eq 'Janelle Schumacher')
#{
	#write-host $name
	#New-MailboxExportRequest -ContentFilter "(Received -gt '$start') -and (Received -lt '$end')" -Mailbox $displayname -ExcludeFolders "#Contacts#" -FilePath $filepath
	#New-MailboxExportRequest -ContentFilter "(Received -gt '$start') -and (Received -lt '$end')" -Mailbox $displayname -FilePath $filepath
	#New-MailboxExportRequest -Name $name -BatchName $batchname -ContentFilter "(Received -gt '$start') -and (Received -lt '$end')" -Mailbox $displayname -FilePath ($exportfolder + $exportid)

	#$count = 0	
	#get-mailboxexportrequest | Select Name, BatchName, Status | Where-Object{$_.Status -eq "InProgress"} | foreach {$count++}		
	#get-mailboxexportrequest | Select Name, BatchName, Status | Where-Object{$_.Status -eq "Queued"} | foreach {$count++}			
	#do
	#{
	#	$count = 0	
	#	get-mailboxexportrequest | Select Name, BatchName, Status | Where-Object{$_.Status -eq "InProgress"} | foreach {$count++}		
	#	get-mailboxexportrequest | Select Name, BatchName, Status | Where-Object{$_.Status -eq "Queued"} | foreach {$count++}
	#	Start-Sleep -s 10
	#}
	#while ($count -ne 0)
	New-MailboxExportRequest -Name $name -BatchName $batchname -ContentFilter "((Received -gt '$start') -and (Received -lt '$end'))" -Mailbox $displayname -FilePath ($exportfolder + $exportid)
		
#}
}

