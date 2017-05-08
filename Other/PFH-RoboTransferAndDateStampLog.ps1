param([String]$from, [string]$to, [string]$roboargs)
$datestamp = (get-date).ToString(‘dd-MM-yyyy_hh-mm-ss’)
$roboargs = $roboargs + " /NP /XF robocopy-backuplog-*.log /LOG+:$from\robocopy-backuplog-"+$datestamp+".log"

write-host $from
write-host $to
write-host $roboargs

Invoke-Expression "robocopy $from $to $roboargs"
