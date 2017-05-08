
' Script to check no jobs are running, shtudown the Batch AOS
' and notify.


  Set objWshShell = Wscript.CreateObject("Wscript.Shell")

  Call Main

  Sub Main

	Call GetAXStoppedUsers("pfg")

      'msg = "Attempting to Shutdown AX Batch AOS. Looking for a shutdown window ( no jobs running )."
      'objWshShell.LogEvent 0, msg
      'Call eMailer(msg, msg, html, attachment1, attachment2, attachment3)

      'Do While BatchJobRunCount() <> 0
      '   Wscript.Sleep 5000
      'Loop

       'msg = "Detected no jobs running, doing shutdown."
       'objWshShell.LogEvent 0, msg
       'Call eMailer(msg, msg, html, attachment1, attachment2, attachment3)
       'Call RestartBatchService
       
  End Sub

  Sub GetAXStoppedUsers(dataareaid)

    'On Error Resume Next

    Set cn = CreateObject("ADODB.Connection")
    cn.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-019\PFWAX" & ";Initial Catalog=" & "PFW_AX2009_Live" & ";Integrated Security=SSPI;"

    Set rs = CreateObject("ADODB.Recordset")
    rs.ActiveConnection = cn
    
    sql = "SELECT ui.NETWORKALIAS FROM dbo.SYSCOMPANYUSERINFO sci INNER JOIN dbo.USERINFO ui ON ui.ID = sci.USERID INNER JOIN " &_
    		"dbo.CUSTTABLE ct ON ct.ACCOUNTNUM = sci.CUSTACCOUNT AND ct.DATAAREAID = sci.DATAAREAID	where ct.BLOCKED = 2 " &_    		
			"and sci.DATAAREAID = '" & dataareaid & "'"

    rs.Open sql,cn,1,1
    rs.MoveFirst

	Do Until rs.EOF
		Wscript.Echo rs(0) & " " & LogUserDistinguishedName(rs(0))
		'If LogUserDistinguishedName(rs(0))
		rs.MoveNext
	Loop

    On Error Goto 0
    err.clear

  End Sub

	Function LogUserDistinguishedName(networkalias)
		'Dim strComputer, strUsername, objWMI, colUsers, objUser
		'On Error Resume Next
		strComputer = "."
		Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\directory\LDAP")
		Set colUsers = objWMI.ExecQuery("SELECT * FROM ds_user where ds_sAMAccountName = '" & networkalias & "'")	   
		If colUsers.Count > 0 Then
			For Each objUser in colUsers	
			'Log("DN:" & objUser.ds_distinguishedName)
				LogUserDistinguishedName = objUser.ds_distinguishedName & " " & objUser.UserAccountDisabled
			Next
		End If
	End Function  

Function eMailer(subject, text, html, attachment1, attachment2, attachment3)

  Set objWshShell = Wscript.CreateObject("Wscript.Shell")
  Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
  Set objNet = Wscript.CreateObject("Wscript.Network")

  '  The mailman object is used for sending and receiving email.
  set mailman = CreateObject("Chilkat.MailMan2")

  '  Any string argument automatically begins the 30-day trial.
  success = mailman.UnlockComponent("MBRRTTMAILQ_WAxTfZe38J6n")
  If (success <> 1) Then
      Return False
      'MsgBox "Component unlock failed"
      WScript.Quit
  End If

  'set vars
  server = "pfnz-srv-015.powerfarming.co.nz"
  mailfrom = "pfnz-srv-020.live-batch-aos@powerfarming.co.nz"
  sendto = "pfwnotify@powerfarming.co.nz"
  subject = "PFNZ-SRV-020:BatchAOSRestart:" & subject

  '  Set the SMTP server.
  mailman.SmtpHost = server

  '  Create a new email object
  set email = CreateObject("Chilkat.Email2")

  email.Subject = subject
  email.Body = text
  If len(html) <> 0 Then
     email.SetHtmlBody html
  End If
  email.From = mailfrom
  email.AddTo "", sendto
  email.AddFileAttachment attachment1
  email.AddFileAttachment attachment2
  email.AddFileAttachment attachment3

  'attachmentContent = "This is the content of a text attachment"

  '  The last argument indicates the charset to use for the attached text.
  '  The string is converted to this charset and attached to the email as a text file attachment.
  '  The charset can be anything: "utf-8", "iso-8859-1", "shift_JIS", "ansi", etc.
  'email.AddStringAttachment2 "myFile.txt",attachmentContent,"ansi"

  success = mailman.SendEmail(email)
  If (success <> 1) Then
   objWshShell.LogEvent 1, "eMailer:Error - " & mailman.LastErrorText
      eMailer = False
  Else
   objWshShell.LogEvent 0, "eMailer:Success - " & " From:" & mailfrom & " Subject:" & subject & " To:" & sendto & " Att:" & attachment1
      eMailer = True
  End If

End Function
