

'**********************************************************************************************
'
'				   Power Farming Group Log Parse Script
'
'**********************************************************************************************

' Author: Michael Barrett
' Date: 06/05/2011

    'Requires: MS LogParser to be installed ( Tested on v2.2 )

    'Set scripting interaction
    'Wscript.Interactive = FALSE

    ' ChangeLog:
    ' 09/05/2011 - MB - Added LogParse Heartbeat Event
    ' 09/05/2011 - MB - Added iCheckpoint for Incremental updates. We can now schedule the events to come in more
    '                   often and only new events are inserted. Hrly?
    ' 10/05/2011 - MB - Added StdOut capture and better logging to NETLOGON for centralisation.
    ' 10/05/2011 - MB - TimeProcessed is now exported per-machine as UTC.
    ' 11/05/2011 - MB - Added clause to get ALL RoboCopyPlus & Backup Exec events

    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")

    'Log Parse HeartBeat Event
    objWshShell.LogEvent 8, "LogParse HeartBeat"

    'Set Start Time
    dtStartNow = Now
    dtLastNow = Now
    
    'Write Init. Log Entries
    log("--> Start")
    Log("File Version: " & ThisScriptModifiedDateStamp())

    'Override this if LogParser is somwhere else.
    exepath = ""

    'Look for LogParser.exe
    If exepath = "" and objFSO.FileExists("C:\Program Files\Log Parser 2.2\LogParser.exe") Then
       exepath = "C:\Program Files\Log Parser 2.2\"
    ElseIf exepath = "" and objFSO.FileExists("C:\Program Files (x86)\Log Parser 2.2\LogParser.exe") Then
       exepath = "C:\Program Files (x86)\Log Parser 2.2\"
    End If

    'Format mth
    if len(month(now-1)) = 1 then
       mth = "0" & CStr(month(now-1))
    else
         mth = month(now-1)
    end  if

    'Format day
    if len(day(now-1)) = 1 then
       dy = "0" & CStr(day(now-1))
    else
         dy = day(now-1)
    end  if

    'Assemble yesterday
    yesterday = year(now-1) & "-" & mth & "-" & dy & " 00:00:00"

    'Assemble cmd
    'cmd = chr(34) & exepath & "LogParser.exe " & chr(34) & " " & chr(34) & "SELECT EventLog, RecordNumber, TO_UTCTIME(TimeGenerated), EventID, EventType, EventTypeName, EventCategory, SourceName, ComputerName, Message FROM System, Application TO Events Where (EventType <> 4 OR EventID = 6013 OR EventID = 6009 OR SourceName = 'RoboCopyPlus' OR SourceName = 'Backup Exec' OR SourceName = 'Backup Exec System Recovery' OR SourceName = 'Microsoft-Windows-WindowsUpdateClient' OR SourceName = 'Windows Update Agent' OR SourceName = 'WSH')" & chr(34) &  " -o:SQL -server:192.168.50.229 -driver:" & chr(34) & "SQL Server" & chr(34) & " -database:itlogcollector -username:logcollector -password:dnaltocs -createtable:ON -iCheckPoint:c:\support\logparse_evtcheck_C.lpc"
	cmd = chr(34) & exepath & "LogParser.exe " & chr(34) & "-i:EVT " & chr(34) & "SELECT EventLog, RecordNumber, TO_UTCTIME(TimeGenerated), EventID, EventType, EventTypeName, EventCategory, SourceName, ComputerName, Message FROM System, Application TO Events" & chr(34) &  " -o:SQL -server:logparserdb.powerfarming.co.nz -driver:" & chr(34) & "SQL Server" & chr(34) & " -database:" & GetComputerName() & " -username:logcollector -password:Dnalt0cs -createtable:ON -iCheckPoint:c:\support\logparse_evtcheck_F.lpc"	
    log("Command To Run: " & cmd)

    'Run cmd
    set retval = objWshShell.Exec (cmd)
      do while retval.Status = 0
           wscript.sleep 100
      loop

    log(">>>>> Unformatted StdOut Starts")
    log(chr(10) & chr(13) & retval.StdOut.ReadAll)
    log(">>>>> Unformatted StdOut Ends")

    Log("Total Time: " & DateDiff("s", dtStartNow, Now) & "secs")
    log("--> End")
    log(" ")

Sub Log(strLog)

	'Clear Err Object
	Err.Clear

	'Enable Error Handling
	On Error Resume Next

	'Open / Create Text File
	'Set fileLogon = objFSO.OpenTextFile("\\powerfarming.co.nz\netlogon\svn-netlogon\LogParser\logs\" & GetComputerName() & "_logparse.txt", 8, True)
	Set fileLogon = objFSO.OpenTextFile("c:\support\" & GetComputerName() & "_logparse.txt", 8, True)	

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