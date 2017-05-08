

Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objNet = Wscript.CreateObject("Wscript.Network")

Call LoopThroughHelpdeskUsers
	
Sub LoopThroughHelpdeskUsers

	'Delete All Users
		Set cn = CreateObject("ADODB.Connection")
		cn.open = "Provider=SQLOLEDB.1; Data Source=" & "SRV-14\SQLEXPRESS" & ";Initial Catalog=" & "iLient" & ";Integrated Security=SSPI;"
		Set rs = CreateObject("ADODB.Recordset")
		rs.ActiveConnection = cn
		sql = "EXECUTE [ilient].[dbo].[DeleteAllSysAidUsers]"
		rs.Open sql,cn,3,3    
		err.clear

	strDomain = "powerfarming.co.nz"
	strGroup = "HelpdeskUser"	
	
	Set objGroup = GetObject("WinNT://" & strDomain & "/" & strGroup)
	Set strMembers = objGroup.Members
	For Each strMember In strMembers		
		If strMember.AccountDisabled = False Then
			Call CreateSysAidHelpdesUser(strMember.Name)
		End IF
	Next
	
End Sub
	
Sub CreateSysAidHelpdesUser(adUserID)	 

	On Error Resume Next

	Dim strComputer, strUsername, objWMI, colUsers, objUser
	strComputer = "."
	Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\directory\LDAP")
	Set colUsers = objWMI.ExecQuery("SELECT * FROM ds_user where ds_sAMAccountName = '" & adUserID & "'")	   
	If colUsers.Count > 0 Then
	   For Each objUser in colUsers	   	  	  
	  
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
			telephoneNumber & "'"
		rs.Open sql,cn,3,3    
		cn.Close
		
	   Next
	End If	 
	
End Sub
