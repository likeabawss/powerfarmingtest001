
	'
	' Script Runs through all server nodes from SysAid and querys them for Windows Updates Status
	'	by Michael Barrett 14/09/2012

    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")

    On Error Resume Next
	
	''Log
	'objWshShell.LogEvent 0, "Pending-Windows-Updates: Started"

    'Set cn = CreateObject("ADODB.Connection")
    'cn.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-028" & ";Initial Catalog=" & "iLient" & ";Integrated Security=SSPI;"
    'Set rs = CreateObject("ADODB.Recordset")
    'rs.ActiveConnection = cn
	'sql = "select c.computer_name from dbo.computer c where c.computer_type = 'Server'"
    'rs.Open sql,cn,1,1
    'rs.MoveFirst
	'Do Until rs.EOF
	'	CheckServerUpdateStatus rs(0)
	'	rs.MoveNext
	'Loop		

	'Log
	objWshShell.LogEvent 0, "Pending-Windows-Updates: Completed"	
	
	CheckServerUpdateStatus "srv-14"
	
	Function CheckServerUpdateStatus( ByVal strServer )

		On Error Goto 0
	
		objWshShell.LogEvent 0, "Pending-Windows-Updates: Processing " & strServer
		
		Dim blnRebootRequired    : blnRebootRequired     = False
		Dim blnRebootPending    : blnRebootPending     = False
		Dim objSession        : Set objSession    = CreateObject("Microsoft.Update.Session", strServer)
		Dim objUpdateSearcher     : Set objUpdateSearcher    = objSession.CreateUpdateSearcher
		Dim objSearchResult    : Set objSearchResult     = objUpdateSearcher.Search(" IsAssigned=1 and IsHidden=0 and Type='Software'")

		Dim i, objUpdate
		Dim intPendingInstalls    : intPendingInstalls     = 0

		For i = 0 To objSearchResult.Updates.Count-1
			Set objUpdate = objSearchResult.Updates.Item(I) 

			If objUpdate.IsInstalled Then
				If objUpdate.RebootRequired Then
					blnRebootPending     = True
				End If
			Else
				intPendingInstalls    = intPendingInstalls + 1
				'If objUpdate.RebootRequired Then    '### This property is FALSE before installation and only set to TRUE after installation to indicate that this patch forced a reboot.
				If objUpdate.InstallationBehavior.RebootBehavior <> 0 Then
					'# http://msdn.microsoft.com/en-us/library/aa386064%28v=VS.85%29.aspx
					'# InstallationBehavior.RebootBehavior = 0    Never reboot
					'# InstallationBehavior.RebootBehavior = 1    Must reboot
					'# InstallationBehavior.RebootBehavior = 2    Can request reboot
					blnRebootRequired     = True
				End If

			End If
		Next

		msg = intPendingInstalls & " updates pending."

		If blnRebootRequired Then
			msg = msg & " Reboot required."
		Else
			msq = msg & " NO Reboot required."
		End If

		If blnRebootPending Then
			msg = msg & " A reboot is waiting to complete a previous install."
		End If 
		
		'Log
		objWshShell.LogEvent 0, "Pending-Windows-Updates: Result for " & strServer & " is: " & msg
		
		'Poke into SysAid DB
		Set cn_ = CreateObject("ADODB.Connection")
		cn_.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-028" & ";Initial Catalog=" & "iLient" & ";Integrated Security=SSPI;"
		Set rs_ = CreateObject("ADODB.Recordset")
		rs_.ActiveConnection = cn_
		sql = "EXECUTE [dbo].[AddWindowsUpdateNote] '" & strServer & "','" & msg & "'"
		rs_.Open sql,cn_,3,3    
		err.clear	
		
	End Function
