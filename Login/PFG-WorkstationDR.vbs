

'**********************************************************************************************
'
'				   Power Farming Group Workstation Backup Script
'
'**********************************************************************************************

'	ChangeLog
'	*********
'	11/04/2013 - MB - Added proc CreateNoBackupMyDocFolder. Also added entry "- **/PersonalNoBackup/**" to file here:
'				\\powerfarming.co.nz\NETLOGON\svn-netlogon\Other\WorkstationDR\incl.txt
'				This informs RSYNC to skip the PersonalNoBackup folder and it's contents anywhere it is found.
'	12/09/2013 - MB - Removed the --delete switch from the RSYNC command executed from the client.
'				As a result, deletions will no longer be synched. --delete-excluded will remain which effectively 
'				deletes those file types that are excluded from the backup.
'				

    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")
	
	Public strComputerName
	Call GetComputerName
	Call Main
	
	Sub Main
				
		On Error Resume Next

		Call CreateNoBackupMyDocFolder
		
		If objFSO.FolderExists("c:\Support") <> True Then	
			objFSO.CreateFolder "c:\Support"
		End If	
		If objFSO.FolderExists("c:\Support\WorkstationDR") <> True Then	
			objFSO.CreateFolder "c:\Support\WorkstationDR"			
		End If			
	
		objFSO.CopyFile "\\powerfarming.co.nz\netlogon\svn-netlogon\other\WorkstationDR\*.*", "c:\Support\WorkstationDR\", True	
		Call CreateBackupBatchCommand
		Call ScheduleHourlyBackup
		
	End Sub
	
	Sub CreateBackupBatchCommand
	
		'Example Command:
		''c:\deltacopy\rsync.exe  -v -rlt -z --delete -m --include-from=/cygdrive/C/support/incl.txt -f 'hide,! */' "/cygdrive/C/Users/" "p0w3r@pfw-workstation-dr.powerfarming.co.nz::NPFW196/Users/"
		
		'ChangeLog:
		'11/03/2013 - MB - Removed the "--delete-excluded" option from the command. Where are system has both C:\USERS and C:\DOCUMENTS AND SETTINGS, when
		'	one was run, it would delete the other, then when the other was run, it would delete the first one's data. It would never finish.
		
		
		On Error Resume Next
		objFSO.DeleteFile("C:\Support\WorkstationDR\ExecuteRsyncWorkstationDR.cmd")
		Set drcmd = objFSO.OpenTextFile("C:\Support\WorkstationDR\ExecuteRsyncWorkstationDR.cmd", 8, True)		
		drcmd.WriteLine "copy \\powerfarming.co.nz\netlogon\svn-netlogon\other\WorkstationDR\incl.txt c:\Support\WorkstationDR\incl.txt /Y"
		drcmd.WriteLine "Start /B /MIN /LOW C:\Support\WorkstationDR\rsync.exe -v -rlt -m -I --size-only " & _
						"--include-from=/cygdrive/C/support/workstationdr/incl.txt -f 'hide,! */' /cygdrive/C/Users/ " & strComputerName & _
						"@pfg-workstation-dr.powerfarming.co.nz::" & strComputerName & "/Users/"
		drcmd.WriteLine "Start /B /MIN /LOW C:\Support\WorkstationDR\rsync.exe -v -rlt -m -I --size-only " & _
						"--include-from=/cygdrive/C/support/workstationdr/incl.txt -f 'hide,! */' " & Chr(34) & "/cygdrive/C/Documents and Settings/" & Chr(34) & " " & strComputerName & _
						"@pfg-workstation-dr.powerfarming.co.nz::" & strComputerName & "/Documents and Settings/"
		drcmd.Close
	
	End Sub
	
	Sub ScheduleHourlyBackup

		'Example Command: 
		'SCHTASKS /Create /RU SYSTEM /SC HOURLY /TN WorkstationDR /TR:C:\SUPPORT\WORKSTATIONDR\EXECUTERSYNCWORKSTATIONDR.CMD /ST 12:36
		
		'Randomize Start Time ( this hour )
		Randomize
		sMin = Cstr(Int((59-0+1)*Rnd+0))
		sHour = Cstr(Hour(Now))		
		If len(sMin) < 2 Then
			sMin = "0" & sMin
		End If
		If len(sHour) < 2 Then
			sHour = "0" & sHour
		End If				
		sTime = Cstr(sHour) & ":" & Cstr(sMin)
		
		On Error Resume Next		
			retval = objWshShell.Run ("SCHTASKS /Delete /TN WorkstationDR /F", 0, TRUE)						
			retval = objWshShell.Run ("SCHTASKS /Delete /TN WorkstationDR-" & objNet.UserName & " /F", 0, TRUE)						
			retval = objWshShell.Run ("SCHTASKS /Create /SC HOURLY /TN WorkstationDR-" & objNet.UserName & " /TR:C:\SUPPORT\WORKSTATIONDR\EXECUTERSYNCWORKSTATIONDR.vbs /ST " & sTime & ":00", 0, TRUE)		
		'retval = objWshShell.Run ("schtasks /Delete /TN WorktstationDR /F", 0, TRUE)				
		'retval = objWshShell.Run ("SCHTASKS /Create /RU SYSTEM /SC HOURLY /TN WorkstationDR /TR:C:\SUPPORT\WORKSTATIONDR\EXECUTERSYNCWORKSTATIONDR.CMD /ST " & sTime & ":00", 0, TRUE)		
	
	End Sub
	
	Sub GetComputerName
		On Error Resume Next
		   strComputerName = objWshShell.RegRead(_
			  "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")
	End Sub	
	
	Sub CreateNoBackupMyDocFolder
	
		Set objWshShell = CreateObject("WScript.Shell")
		myDocPath = objWshShell.SpecialFolders("MyDocuments")
		If objFSO.FolderExists(myDocPath & "\PersonalNoBackup") <> True Then			
			objFSO.CreateFolder myDocPath & "\PersonalNoBackup"
		End If
	
	End Sub		