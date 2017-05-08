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

$server = "PFNZ-SRV-030"
$start = '12/01/2011'
$end = '01/01/2012'

#Loop through each mailbox
foreach ($mailbox in (Get-Mailbox -server $server)) {

$displayname = $mailbox.Displayname
$filepath = "\\pfnz-srv-030\archives\" + $displayname + ".pst"	

#if ($displayname -eq 'Michael Barrett')
#{
	write-host $displayname
	New-MailboxExportRequest -ContentFilter "(Received -gt '$start') -and (Received -lt '$end')" -Mailbox $displayname -FilePath $filepath
#}
}