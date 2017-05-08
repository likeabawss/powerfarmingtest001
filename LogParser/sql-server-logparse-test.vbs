

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

    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")

    objWshShell.LogEvent 8, "LogParse HeatBeat"

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
    cmd = chr(34) & exepath & "LogParser.exe " & chr(34) & " " & chr(34) & "SELECT EventLog, RecordNumber, TimeGenerated, EventID, EventType, EventTypeName, EventCategory, SourceName, ComputerName, Message FROM System, Application TO EventsTesting Where (EventType <> 4 OR EventID = 6013 OR EventID = 6009) and TimeWritten > '" _
                     & yesterday & "'" & chr(34) & " -o:SQL -server:PFNZ-SRV-029 -driver:" & chr(34) & "SQL Server" & chr(34) & " -database:itlogcollector -username:logcollector -password:dnaltocs -createtable:ON -iCheckPoint:c:\temp\logparse_evtcheck.lpc"

    'wscript.echo cmd

    'Run cmd
    retval = objWshShell.Run (cmd, 0, FALSE)
