On Error Resume Next

    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")	

	objWshShell.RegDelete "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\"

	'Restart WSUS service
	strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
	strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
	strUpdate = objWshShell.Run ("wuauclt /detectnow", 0, TRUE)
