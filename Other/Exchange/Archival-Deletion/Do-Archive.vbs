
'Dim obj vars
Public objWshShell		'WSH Shell
Public objFSO			'WSH FileSystemObject
Public objNET			'WSH Networking Object

'Load Intrinsic Objects
Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objNet = Wscript.CreateObject("Wscript.Network")
Set args = WScript.Arguments

    'Set Start Time
    dtStartNow = Now
    dtLastNow = Now

Call Main

Sub Main
    
    'Init.
	Dim taskName, taskCmd, retval
    log("--> Start")
	Log("Script Path: " & Wscript.ScriptName)
    log("Script Version: " & ThisScriptModifiedDateStamp())	

	'Check Args
	If args.Count = 0 then
		log(UCase(Wscript.ScriptName) & " was run without any arguments, please supply a Server, Month Offset, Exchange Store, and PST location. Quitting.")	
		
		'objWshShell.LogEvent 1, "<ScriptLogging>" & _
		'						"<ScriptName>" & Wscript.ScriptName & "</ScriptName>" & _
		'						"<FileModDateStamp>" & ThisScriptModifiedDateStamp() & "</FileModDateStamp>" & _
		'						"<TaskName>" & "ERROR" & "</TaskName>" & _
		'						"<TaskCmd>" & "NOARGS" & "</TaskCmd>" & _
		'						"<ProcTimeSecs>" & "0" & "</ProcTimeSecs>"  & _		
		'						"<OutCmd>" & Wscript.ScriptName & " was run without any arguments, please supply a scheduled id and retry!" & "</OutCmd>" & _
		'						"</ScriptLogging>"
		'objWshShell.Popup UCase(Wscript.ScriptName) & " was run without any arguments, please supply a schedule/section ID and retry!", 10, UCase(Wscript.ScriptName) & " - NOARGS"		
		log("Total Time: " & DateDiff("s", dtStartNow, Now) & "secs")
		log("--> End")
		log(" ")		
		Wscript.Quit
	Else
		'Dump Arguments
		For Each arg in args
			log("	Arg: " & arg)
		Next 			
		''log("Scheduled / Section Argument: " & args.Item(0))
	End If

	'Store Current Locale
	origLocale = objWshShell.RegRead("HKCU\Control Panel\International\LocaleName")
	log("Locale Before Change:" & objWshShell.RegRead("HKCU\Control Panel\International\LocaleName"))

	'Detect Current Locale and Appply US.
	Select Case origLocale
		Case "en-NZ"
			objWshShell.Run "Regedit /S " & "\\powerfarming.co.nz\NETLOGON\svn-netlogon\Other\Exchange\Archival-Deletion\HKCU_en-US.reg", 7, TRUE
		Case "en-AU"
			objWshShell.Run "Regedit /S " & "\\powerfarming.co.nz\NETLOGON\svn-netlogon\Other\Exchange\Archival-Deletion\HKCU_en-US.reg", 7, TRUE		
		Case Else
			'Only NZ and AU locales permitted
			log("Only NZ and AU locales permitted.")			
			log("<-- End")			
			Wscript.Quit
	End Select

	log("Locale During Archive Processing:" & objWshShell.RegRead("HKCU\Control Panel\International\LocaleName"))
	
	'Detect Current Locale and Appply back to Original
	Select Case origLocale
		Case "en-NZ"
			objWshShell.Run "Regedit /S " & "\\powerfarming.co.nz\NETLOGON\svn-netlogon\Other\Exchange\Archival-Deletion\HKCU_en-NZ.reg", 7, TRUE
		Case "en-AU"
			objWshShell.Run "Regedit /S " & "\\powerfarming.co.nz\NETLOGON\svn-netlogon\Other\Exchange\Archival-Deletion\HKCU_en-AU.reg", 7, TRUE		
	End Select	
	
	log("Locale After Processing:" & objWshShell.RegRead("HKCU\Control Panel\International\LocaleName"))		
		
	log("<-- End")			
	Wscript.Quit		
		
End Sub

Function ThisScriptModifiedDateStamp()
  On Error Resume Next
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFSO.GetFile(Wscript.ScriptFullName)
  ThisScriptModifiedDateStamp = CTimeStamp(CDate(objFile.DateLastModified))
  On Error Goto 0
End Function

Function GetComputerName
    Err.Clear
    On Error Resume Next

	   'Get computer name from registry
	   GetComputerName = objWshShell.RegRead(_
	      "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")

    On Error Goto 0
End Function  

Function CTimeStamp(dt)
	Err.Clear 
	On Error Resume Next

	'Set Dates/Times
	strMonth = Month(dt)
	strDay = Day(dt)
	strHour = Hour(dt)
	strMin = Minute(dt)		
	strSec = Second(dt)
	
	'Update Parameter Length for standardization
	If strMonth < 10 Then
		strMonth = "0" & strMonth	
	End If
	If strDay < 10 Then
		strDay = "0" & strDay
	End If
	If strHour < 10 Then
		strDay = "0" & strDay
	End If
	If strMin < 10 Then
		strMin = "0" & strMin
	End If		
	If strSec < 10 Then
		strSec = "0" & strSec
	End If			
	
	'Current Time Stamp
	CTimeStamp = Year(Now) & strMonth & strDay & strHour & strMin & strSec
	
	On Error Goto 0
End Function

Sub Log(strLog)
	Err.Clear
	On Error Resume Next

	'Open / Create Text File
	Set fileLogon = objFSO.OpenTextFile("\\powerfarming.co.nz\netlogon\svn-netlogon\Other\Exchange\Archival-Deletion\Logs\" & GetComputerName() & "_Exch-Auto-Archive.log", 8, True)
	'Set fileLogon = objFSO.OpenTextFile("c:\Support\" & GetComputerName() & "_Exch-Auto-Archive.log", 8, True)

	'Get seconds since last log
	dtJustNow = Now
	sLastLog = DateDiff("s", dtLastNow, dtJustNow)
	
	'Write
	fileLogon.Write dtJustNow & " (" & sLastLog & "secs) " & strLog
	fileLogon.WriteLine
	dtLastNow = Now

  'Close
  fileLogon.Close

  On Error Goto 0
  Err.Clear
End Sub  