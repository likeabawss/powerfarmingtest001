
'**********************************************************************************************
'
'				   Power Farming Group Auto SVN Script
'
'**********************************************************************************************

' Author: Michael Barrett
' Date: 06/05/2011

  'used slik-svn from here: http://www.sliksvn.com/ to SVN command-line tool that would work.
  'installed and ran TortoiseSVN client from server, browsed repo so it would prompt and save
  'svn userid and password as adminjobuser.

  'Change Log:
  '***********
  '01/06/2011 - MEB - Added logging to central replicated folder 
  '02/06/2011 - MEB - Centralised script to and added tasks lists.
  '02/06/2011 - MEB - Added logging of tasks STDOUT to event log.
  '03/06/2011 - MEB - Added Argument as Schedule / Section ID.
   
   'on error resume next
	Set objWshShell = CreateObject("Wscript.Shell")
	Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
	Set objNet = Wscript.CreateObject("Wscript.Network")
	Set args = WScript.Arguments
		'First Argument is the Schedule
			
    'Set Start Time
    dtStartNow = Now
    dtLastNow = Now
    
    'Init.
	Dim taskName, taskCmd, retval
    log("--> Start")
	Log("Script Path: " & Wscript.ScriptName)
    log("Script Version: " & ThisScriptModifiedDateStamp())	
		
	'Check Args
	If args.Count = 0 then
		log(UCase(Wscript.ScriptName) & " was run without any arguments, please supply a schedule/section ID and retry! Quitting.")	
		objWshShell.LogEvent 1, "<ScriptLogging>" & _
								"<ScriptName>" & Wscript.ScriptName & "</ScriptName>" & _
								"<FileModDateStamp>" & ThisScriptModifiedDateStamp() & "</FileModDateStamp>" & _
								"<TaskName>" & "ERROR" & "</TaskName>" & _
								"<TaskCmd>" & "NOARGS" & "</TaskCmd>" & _
								"<ProcTimeSecs>" & "0" & "</ProcTimeSecs>"  & _		
								"<OutCmd>" & Wscript.ScriptName & " was run without any arguments, pleas supply a scheduled id and retry!" & "</OutCmd>" & _
								"</ScriptLogging>"
		objWshShell.Popup UCase(Wscript.ScriptName) & " was run without any arguments, please supply a schedule/section ID and retry!", 10, UCase(Wscript.ScriptName) & " - NOARGS"		
		log("Total Time: " & DateDiff("s", dtStartNow, Now) & "secs")
		log("--> End")
		log(" ")		
		Wscript.Quit
	Else
		log("Scheduled / Section Argument: " & args.Item(0))
	End if		
	
	'Open task file for this machine and run and log tasks.
	Set taskfile = objFSO.OpenTextFile("\\powerfarming.co.nz\NETLOGON\svn-netlogon\AutoSvn\" & trim(GetComputerName()) & "_autosvn_tasks.txt"  ,1)
	Do While taskfile.AtEndOfStream <> True
		cLine = taskfile.ReadLine
		
		'Skiplines until schedule
		If ScheduleFound = FALSE and cLine = "[SCHEDULE:" & trim(UCase(args.Item(0))) & "]" then
			ScheduleFound = TRUE
			cLine = taskfile.ReadLine			
		End If
		
		If ScheduleFound = TRUE Then	
			'halt
			If UCase(Trim(cLine)) = "STOP" Then			
				log("STOP")
				Exit Do
			End If
			
			If Instr(cLine, "|") <> 0 Then
				taskName = mid(cLine, 1, Instr(cLine, "|") - 1)
				taskCmd = mid(cLine, Instr(cLine, "|") + 1)
				
				log("Command Name: " & taskName)
				log("Command To Run: " & taskCmd)
				lgStartNow = Now
				
					Err.Clear
					On Error Resume Next
				set retval = objWshShell.Exec (taskCmd)
					If Err.Number <> 0 Then
						stdout = err.number & "|" & err.description
					End If
				do while retval.Status = 0
				   wscript.sleep 100
				loop
				log(">>>>> Unformatted StdOut Starts")
				stdout = retval.StdOut.ReadAll
				log(chr(10) & chr(13) & "                            " & stdout)
				log(">>>>> Unformatted StdOut Ends")	

				'Log Results to Event Log
				objWshShell.LogEvent 4, "<ScriptLogging>" & _
										"<ScriptName>" & Wscript.ScriptName & "</ScriptName>" & _
										"<FileModDateStamp>" & ThisScriptModifiedDateStamp() & "</FileModDateStamp>" & _
										"<TaskName>" & taskName & "</TaskName>" & _
										"<TaskCmd>" & taskCmd & "</TaskCmd>" & _
										"<ProcTimeSecs>" & DateDiff("s", lgStartNow, Now) & "</ProcTimeSecs>"  & _		
										"<OutCmd>" & stdout & "</OutCmd>" & _
										"</ScriptLogging>"
			End If
		End If
	Loop	

    log("Total Time: " & DateDiff("s", dtStartNow, Now) & "secs")
    log("--> End")
    log(" ")	
	  

Sub Log(strLog)
	Err.Clear
	On Error Resume Next

	'Open / Create Text File
	'Set fileLogon = objFSO.OpenTextFile("\\powerfarming.co.nz\NETLOGON\svn-netlogon\AutoSvn\Logs\" & GetComputerName() & "_autosvn.log", 8, True)
	Set fileLogon = objFSO.OpenTextFile("c:\Support\" & GetComputerName() & "_autosvn.log", 8, True)

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