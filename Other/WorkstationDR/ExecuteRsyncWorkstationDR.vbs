
    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")	
	retval = objWshShell.Run ("cmd /c C:\SUPPORT\WORKSTATIONDR\ExecuteRsyncWorkstationDR.cmd", 0, TRUE)