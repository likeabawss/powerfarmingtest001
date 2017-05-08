
Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objNet = Wscript.CreateObject("Wscript.Network")

Call LoopThroughHelpdeskUsers
	
Sub LoopThroughHelpdeskUsers

	'Disable All Users in SysAid ( not Admins )
		Set cn = CreateObject("ADODB.Connection")
		cn.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-028" & ";Initial Catalog=" & "iLient" & ";Integrated Security=SSPI;"
		Set rs = CreateObject("ADODB.Recordset")
		rs.ActiveConnection = cn
		'sql = "EXECUTE [ilient].[dbo].[DeleteAllSysAidUsers]"
		
		sql = "update usr " &_
				"set usr.[disable] = 'Y' " &_
				"from [dbo].[sysaid_user] usr " &_				
				"where usr.administrator = 'N' "
				
		rs.Open sql,cn,3,3    
		err.clear	
	
	strDomain = "powerfarming.co.nz"
	strNTLMDomain = "POWER"
	strGroup = "HelpdeskUser"	

	Set objGroup = GetObject("WinNT://" & strDomain & "/" & strGroup)
	Set strMembers = objGroup.Members
	For Each strMember In strMembers		
		If strMember.AccountDisabled = False Then
			Call EnableSysAidHelpdeskUser(strNTLMDomain & "\" & strMember.Name)			
		End If
	Next
	
	'Find and Disable all users without email addresses.
	sql = "update [ilient].[dbo].[sysaid_user] " & _
			"set [disable] = 'Y'" & _
			"where (email_address = '' OR email_address is null)"
			
	rs.Open sql,cn,3,3    
	err.clear	
	
	'Find and Delete all users who are disabled.
	'SysAid seems to want to use it's own rules for disabling and enabling users which means at odd times enabled users are 
	'toggeld and vice versa. Not idea.
	'We'll rely on the scheduled job run at 1am each morning to bring new accounts back, and this script to clean up after it.
	sql = "delete from [ilient].[dbo].[sysaid_user] " & _
			"where [disable] = 'Y'"
			
	rs.Open sql,cn,3,3    
	err.clear		
	
	End Sub

Sub EnableSysAidHelpdeskUser(adUserID)

	'Disable All Users in SysAid ( not Admins )
		Set cn = CreateObject("ADODB.Connection")
		cn.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-028" & ";Initial Catalog=" & "iLient" & ";Integrated Security=SSPI;"
		Set rs = CreateObject("ADODB.Recordset")
		rs.ActiveConnection = cn		
		
		sql = "update usr " &_
				"set usr.[disable] = 'N' " &_
				"from [dbo].[sysaid_user] usr " &_				
				"where usr.user_name = '" & adUserID & "'"
				
		rs.Open sql,cn,3,3    
		err.clear	

End Sub
	
Sub CreateSysAidHelpdesUser(adUserID)	 

	On Error Resume Next

	Dim strComputer, strUsername, objWMI, colUsers, objUser
	strComputer = "."
	Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\directory\LDAP")
	Set colUsers = objWMI.ExecQuery("SELECT * FROM ds_user where ds_sAMAccountName = '" & adUserID & "'")	   
	If colUsers.Count > 0 Then
	   For Each objUser in colUsers	   	  	  
	  
		'Wscript.Echo adUserID & ", " & objUser.ds_company
		'Wscript.Echo adUserID & ", " & GetCompanyID(objUser.ds_company)
					
		If Trim(objUser.ds_telephoneNumber) <> "" Then
			telephoneNumber = objUser.ds_telephoneNumber & " extn. 7" & Right(objUser.ds_telephoneNumber,3)			
		Else	
			telephoneNumber	= ""
		End If				

		Set cn = CreateObject("ADODB.Connection")
		cn.open = "Provider=SQLOLEDB.1; Data Source=" & "SRV-14\SQLEXPRESS" & ";Initial Catalog=" & "iLient" & ";Integrated Security=SSPI;"
		Set rs = CreateObject("ADODB.Recordset")
		rs.ActiveConnection = cn
		sql = "EXECUTE [ilient].[dbo].[AddRecreateSysAidUser] '" &_ 
			adUserID & "','" &_
			objUser.ds_mail & "','" &_
			objUser.ds_givenName & "','" &_
			Replace(objUser.ds_sn,"'","") & "','" &_
			objUser.ds_mobile & "','" &_
			telephoneNumber & "'," &_
			GetCompanyID(objUser.ds_company)
		rs.Open sql,cn,3,3    
		cn.Close
		
	   Next
	End If	 
	
End Sub

Function GetCompanyID(companyName)

    On Error Resume Next

    Set cn = CreateObject("ADODB.Connection")
    cn.open = "Provider=SQLOLEDB.1; Data Source=" & "SRV-14\SQLEXPRESS" & ";Initial Catalog=" & "ILIENT" & ";Integrated Security=SSPI;"

    Set rs = CreateObject("ADODB.Recordset")
    rs.ActiveConnection = cn
    sql = "select isnull([dbo].[fn_GetCompanyID]('" & companyName & "'),7)"
    sql
    rs.Open sql,cn,1,1
    rs.MoveFirst

    GetCompanyID = rs(0)

    On Error Goto 0
    err.clear

End Function
