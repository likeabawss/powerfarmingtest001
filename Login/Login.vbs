'Load Intrinsic Objects
Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objNet = Wscript.CreateObject("Wscript.Network")

'	ChangeLog:
'	**********
'	25/07/2011 - MB - Added RemoveNonLocalNetworkPrinters. This will attempt to remove printers from the local session that are non in the local subnet.

'Run Script
'objWshShell.Run "wscript " & "\\powerfarming.co.nz\netlogon\svn-netlogon\login\RemoveNonLocalNetworkPrinters.vbs", 7, FALSE
objWshShell.Run "wscript " & "\\powerfarming.co.nz\netlogon\svn-netlogon\login\BaseLogin.wsf", 7, FALSE
