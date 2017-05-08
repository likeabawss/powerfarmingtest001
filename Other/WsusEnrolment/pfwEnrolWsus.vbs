


On Error Resume Next

    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")	
	
		'Get Configured WSUS Server
		strWSUSServer = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")
		strWSUSServer = Trim(strWSUSServer)
		' = strWSUSServer
		'If strWSUSServer <> "http://PFNZ-SRV-028.powerfarming.co.nz:8530" AND strWSUSServer <> "" Then				
		
			'Re-Home Client
			If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFNZ-SRV-028.powerfarming.co.nz:8530" Then
			   'Log
			   Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Re-homing WSUS client " & UCase(strComputerName) & " to WSUS server PFNZ-SRV-028.powerfarming.co.nz.")
					   'Install Req. Keys
			   objWshShell.ReqWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate", 0, "REG_BINARY"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFNZ-SRV-028.powerfarming.co.nz:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFNZ-SRV-028.powerfarming.co.nz:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 2, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
			   'Restart WSUS service
			   strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("wuauclt /detectnow", 0, TRUE)
			'End If
			
		End If