
	Set objWshShell = Wscript.CreateObject("Wscript.Shell")
	Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
	Set objNet = Wscript.CreateObject("Wscript.Network")
			
	procname = Wscript.Arguments(0)
	If trim(procname) = "" then
		Wscript.Quit
	Else	
		Call KillProc(procname)
	End If	

	Sub KillProc(proc)

	'Allow the killing of other users' priveleges.
	Set objLoc = createobject("wbemscripting.swbemlocator")
	objLoc.Security_.privileges.addasstring "sedebugprivilege", true
	
		'Clear Error Object
		Err.Clear

		'Disable Error Handling
		'On Error Resume Next

		strComputer = "."
		'Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate, (Debug)}!\\" & strComputer & "\root\cimv2")
		Set colProcesses = GetObject("winmgmts:" & _
		   "{impersonationLevel=impersonate, (Debug)}!\\" & strComputer & _
		   "\root\cimv2").ExecQuery("Select * from Win32_Process")

			For Each objProcess in colProcesses
				If Ucase(objProcess.Description) = Ucase(proc) Then
				   retval = objProcess.Terminate
				End If
			Next

		'Disable Error Handling
		On Error Goto 0

	End Sub
