
'
'	ChangeLog:
'	**********
'	09/02/2015 - MB - Added Office Location Contact Attribute
'	10/02/2015 - MB - Allowed Disabled accounts through, so we can filter them out of Outlook views.
'	10/02/2015 - MB - Updated to update new field ADDisabled on form to exclude from views.
'	11/02/2015 - MB - Large changes have been made to this script to attempt form upgrades where possible
'						and recreate contacts from AD where not.
'	12/02/2015 - MB - Added Division.
'	24/02/2015 - MB - Added Brands for Salesmen
'	25/02/2015 - MB - Added error checking, recreation of contact if update failure.
'

Const ADS_PROPERTY_UPDATE = 2
Dim intLatestFormVersion

Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    'Set Start Time
    dtStartNow = Now
    dtLastNow = Now
    Log("**GroupStaff-PublicFolder-ContactLoader Starts**")
strFolderPath = "Group Staff"
Call KillOutlook()
Call WaitForOutlook()
Log("**GroupStaff-PublicFolder-ContactLoader Finishes**")

'Iterate through Query
Sub IterateSQLContacts
	On Error Resume Next
	'On Error Goto 0
	intCounter = 0
	Set cn = CreateObject("ADODB.Connection")
	cn.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-019\PFWAX" & ";Initial Catalog=" & "DataMart" & ";Integrated Security=SSPI;"
	Set rs = CreateObject("ADODB.Recordset")
	rs.ActiveConnection = cn
	sql = "Select * FROM [DataMart].[dbo].[SYS - AD Active Group Staff]"
	rs.Open sql,cn,1,1
		
	'Run Update / Create
	If rs.EOF = False Then					
		Do While Not rs.EOF
				
			Call exchGalHide(rs, "FALSE")
					
			intCounter = intCounter + 1
			If ContactExists(rs("objectGUID")) = True Then
				Log(intCounter & " - Processing Contact:" & rs("objectGUID") & "(" & rs("FullName") & ")")
				Call Update(rs)
			Else
				Log(intCounter & " - Create Contact:" & rs("objectGUID") & "(" & rs("FullName") & ")")
				Call CreateNew(rs)
			End If	
			rs.MoveNext		
		Loop	
	End If
	Log(intCounter & " contacts processed.")	
End Sub

Sub exchGalHide(rs, status)

			'There appears to be no way to detect a NULL value using ADSI
			On Error Resume Next
			Set galHide = GetObject("LDAP://" & rs("distinguishedName")) 
			galHideValue = ""
			galHideValue = UCase(CStr(galHide.Get("msExchHideFromAddressLists")))
			If galHideValue = "" Then
				galHideValue = "FALSE"
			End If
			On Error Goto 0
		
			'WScript.Echo rs("FullName") & " - status: " & status & " - galHideValue:" & galHideValue
		
			If UCase(status) <> UCase(galHideValue) Then
				galHide.Put "msExchHideFromAddressLists", UCase(status)
				galHide.SetInfo														
				Log("X - Set to Hide (" & UCase(status) & ") Contact in GAL:" & rs("objectGUID") & "(" & rs("FullName") & ")")						
			End If
		
			Set galHide = Nothing

End Sub

Sub DeleteBogusContacts

	On Error Goto 0	
	Set fldr = GetPublicFolder(strFolderPath)
	Set itms = fldr.Items
	Set itm = itms.Find("([ADobjectGUID] = " & Chr(34) & "formcheck" & Chr(34) & " OR [ADobjectGUID] = " & Chr(34) & Chr(34) & ")")	
	guidCheck = ""
	Err.Clear
	Do While IsEmpty(itm) = False		
		itm.Delete
		Log("X - Deleted Folder Contact:" & itm.UserProperties.Find("ADobjectGUID").Value & " (" & itm.FullName & ")")
		Set itm = itms.FindNext
	Loop

End Sub

Sub DeleteNonRequiredSQLContacts
	On Error Resume Next
	intCounter = 0
	Log("Deleting Non-required Contacts from Public Folder and Hiding same from Exchange GAL.")							
	Set cn = CreateObject("ADODB.Connection")
	cn.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-019\PFWAX" & ";Initial Catalog=" & "DataMart" & ";Integrated Security=SSPI;"
	Set rs = CreateObject("ADODB.Recordset")
	rs.ActiveConnection = cn
	sql = "SELECT au.* FROM [DataMart].[dbo].[SYS - AD ALL Users] au WHERE not exists(select * from [DataMart].[dbo].[SYS - AD Active Group Staff] act where au.objectGUID = act.objectGUID)"
	rs.Open sql,cn,1,1

	Set fldr = GetPublicFolder(strFolderPath)
	Set itms = fldr.Items
	
	'Run Update / Create
	If rs.EOF = False Then					
		Do While Not rs.EOF			
			
			Call exchGalHide(rs, "TRUE")			

			Set itm = itms.Find("[ADobjectGUID] = " & Chr(34) & rs("objectGUID") & Chr(34))
			nulchk = ""
			nulchk = itm.UserProperties.Find("ADobjectGUID").Value
			If nulchk <> "" Then
				'WScript.Echo "Deleted " & rs("distinguishedName")
				Log("X - Deleted Contact:" & rs("objectGUID") & "(" & rs("FullName") & ")")
				itm.Delete
				Set itm = Nothing
			End If				
						
			rs.MoveNext		
		Loop	
	End If
	Log("Deletion / Hiding Checking Complete.")							
End Sub

Sub SetDivision(rs)

	On Error Resume Next

	Set objUser = GetObject("LDAP://" & rs("distinguishedName"))
	If InStr(UCase(rs("distinguishedName")), "NEW ZEALAND RETAIL") <> 0 Then
		If rs("division") <> "Retail" Then
			objUser.Put "division", "Retail"
			objUser.SetInfo
		End If			
	Else
		If rs("division") <> "Wholesale" Then
			objUser.Put "division", "Wholesale"
			objUser.SetInfo
		End If			
	End If	
			

End Sub

' Checks for the existence of a contact using the AD GUID. Also checks that the formversion of the default form in the folder
' matches the contact. If it doesn't an attempt is made to upgrade the contact, however if this fails, it is deleted and a 
' new one generated fresh from AD data. The FormVersion field is set for every new contact, to the latest version of the Form.
Function ContactExists(objectGUID)
	'On Error Resume Next
	On Error Goto 0
	Set fldr = GetPublicFolder(strFolderPath)	
	Set itms = fldr.Items

	'Attempt to Find Contact		
	Set itm = itms.Find("[ADobjectGUID] = " & Chr(34) & objectGUID & Chr(34))				
	If itm Is Nothing Then
		ContactExists = False
	Else
		
		'Contact was found so Open and Close the Item to Attempt Form Upgrade
		Set itm_ = itms.Find("[ADobjectGUID] = " & Chr(34) & objectGUID & Chr(34))
		itm_.Save
		itm_.Close(0)	
	
		'The Version Label Caption is the Most Reliable way to detect whether
		'the upgrade worked. Typically only contacts that have had no changes
		'since the default form was changed are upgradeable.
		Set objGeneralPage = itm.GetInspector.ModifiedFormPages("General")
		Set FormVersion = objGeneralPage.Controls("version")
		itm.UserProperties.Find("FormVersion").Value = FormVersion.Caption
		itm.Save
	
		'Check if the upgrade succeeded.
		If itm.UserProperties.Find("FormVersion").Value = intLatestFormVersion Then
			ContactExists = True
		Else
			'Upgrade failed. Delete record, ready for recreation from AD.
			itm.Delete
			ContactExists = False
		End If
	End If	
End Function

Sub Update(rs)

	On Error Resume Next
	Set fldr = GetPublicFolder(strFolderPath)
	Set itms = fldr.Items
	Set itm = itms.Find("[ADobjectGUID] = " & Chr(34) & rs("objectGUID") & Chr(34))
	If itm Is Nothing Then
	   Exit Sub
	Else								
		If itm.UserProperties.Find("ADuSNChanged").Value <> rs("uSNChanged") Then
		
		Log("    	Updated Contact:" & rs("objectGUID") & "(" & rs("FullName") & ")")
		
			'Standard
			itm.FileAs = rs("FullName")
			itm.FullName = rs("FullName")
			itm.BusinessTelephoneNumber = rs("TelephoneNumber")
			itm.MobileTelephoneNumber = rs("MobileNumber")
			itm.BusinessFaxNumber = rs("FaxNumber")
			itm.Email1Address = rs("Email")
			itm.CompanyName = rs("Company")
			itm.JobTitle = rs("JobTitle")
			itm.OfficeLocation = rs("Office")			
			itm.UserProperties.Find("ADOffice").Value = rs("Office")
			
			'UserProperties
			itm.UserProperties.Find("ADSAMAccountName").Value = rs("SAMAccountName")
			itm.UserProperties.Find("ADFullName").Value = rs("FullName")
			itm.UserProperties.Find("ADTelephoneNumber").Value = rs("TelephoneNumber")
			itm.UserProperties.Find("ADFirstName").Value = rs("FirstName")
			itm.UserProperties.Find("ADSurname").Value = rs("Surname")
			itm.UserProperties.Find("ADFaxNumber").Value = rs("FaxNumber")
			itm.UserProperties.Find("ADAccountStatus").Value = rs("ADAccountStatus")
			itm.UserProperties.Find("ADEmail").Value = rs("Email")
			itm.UserProperties.Find("ADDepartment").Value = rs("Department")
			itm.UserProperties.Find("LocalDepartment").Value = rs("Department")						
			itm.UserProperties.Find("ADCompany").Value = rs("Company")
			itm.UserProperties.Find("ADManager").Value = rs("Manager")
			itm.UserProperties.Find("ADMobileNumber").Value = rs("MobileNumber")
			itm.UserProperties.Find("ADdistinguishedName").Value = rs("distinguishedName")
			itm.UserProperties.Find("ADdistinguishedContainer").Value = rs("distinguishedContainer")
			itm.UserProperties.Find("ADuSNChanged").Value = rs("uSNChanged")
			itm.UserProperties.Find("ADobjectGUID").Value = rs("objectGUID")
			itm.UserProperties.Find("ADDeptManager").Value = rs("DeptManager")
			itm.UserProperties.Find("DeptManager").Value = rs("DeptManager")
			itm.UserProperties.Find("ADAccountDisabled").Value = rs("Disabled")
			
			itm.UserProperties.Find("Division").Value = rs("division")
			itm.UserProperties.Find("ADDivision").Value = rs("division")
			
			itm.UserProperties.Find("NonPerson").Value = rs("NonPerson")
			itm.UserProperties.Find("ADNonPerson").Value = rs("NonPerson")
			
			itm.UserProperties.Find("PendingInactive").Value = rs("PendingInactive")
			itm.UserProperties.Find("ADPendingInactive").Value = rs("PendingInactive")
						
			itm.UserProperties.Find("SpeedDial").Value = rs("SpeedDial")			
			itm.UserProperties.Find("ADSpeedDial").Value = rs("SpeedDial")
			
			itm.Business2TelephoneNumber = rs("TelephoneExtn")
			itm.UserProperties.Find("ADTelephoneExtn").Value = rs("TelephoneExtn")
			
			itm.UserProperties.Find("CreatedStatus").Value = rs("CreatedStatus")	
			
			If rs("SalesmenBrands") <> "" Then
				arSalesmenBrands = Split(rs("SalesmenBrands"), ",")
				For Each brand In arSalesmenBrands
					itm.UserProperties.Find("rep" & brand).Value = -1
				Next
			End If			
			
			itm.Save
			itm.Close(0)
			
		End If
	End If		
	
	If Err.Number <> 0 Then
		Log(intCounter & " - Error Updating Contact:" & rs("objectGUID") & "(" & rs("FullName") & ")" & " - " & Err.Number & " - " & Err.Description)		
		Log(intCounter & " - Deleting Contact:" & rs("objectGUID") & "(" & rs("FullName") & ")")
		itm.Delete		
		CreateNew(rs)					
	Else 
		itm.Close(0)				
	End If	
	
End Sub

Sub CreateNew(rs)

	On Error Resume Next

	Set itm = GetPublicFolder(strFolderPath).Items.Add("IPM.Contact.InternalGlobalContact")
	
	'Standard
	itm.FileAs = rs("FullName")
	itm.FullName = rs("FullName")
	itm.BusinessTelephoneNumber = rs("TelephoneNumber")
	itm.MobileTelephoneNumber = rs("MobileNumber")
	itm.BusinessFaxNumber = rs("FaxNumber")
	itm.Email1Address = rs("Email")
	itm.CompanyName = rs("Company")
	itm.JobTitle = rs("JobTitle")
	itm.OfficeLocation = rs("Office")
	itm.UserProperties.Find("ADOffice").Value = rs("Office")			
	
	'UserProperties
	itm.UserProperties.Find("ADSAMAccountName").Value = rs("SAMAccountName")
	itm.UserProperties.Find("ADFullName").Value = rs("FullName")
	itm.UserProperties.Find("ADTelephoneNumber").Value = rs("TelephoneNumber")
	itm.UserProperties.Find("ADFirstName").Value = rs("FirstName")
	itm.UserProperties.Find("ADSurname").Value = rs("Surname")
	itm.UserProperties.Find("ADFaxNumber").Value = rs("FaxNumber")
	itm.UserProperties.Find("ADAccountStatus").Value = rs("ADAccountStatus")
	itm.UserProperties.Find("ADEmail").Value = rs("Email")
	itm.UserProperties.Find("ADDepartment").Value = rs("Department")
	itm.UserProperties.Find("LocalDepartment").Value = rs("Department")
	itm.UserProperties.Find("ADCompany").Value = rs("Company")
	itm.UserProperties.Find("ADManager").Value = rs("Manager")
	itm.UserProperties.Find("ADMobileNumber").Value = rs("MobileNumber")
	itm.UserProperties.Find("ADdistinguishedName").Value = rs("distinguishedName")
	itm.UserProperties.Find("ADdistinguishedContainer").Value = rs("distinguishedContainer")
	itm.UserProperties.Find("ADuSNChanged").Value = rs("uSNChanged")
	itm.UserProperties.Find("ADobjectGUID").Value = rs("objectGUID")
	itm.UserProperties.Find("ADDeptManager").Value = rs("DeptManager")
	itm.UserProperties.Find("DeptManager").Value = rs("DeptManager")
	itm.UserProperties.Find("ADAccountDisabled").Value = rs("Disabled")	
	itm.UserProperties.Find("FormVersion").Value = intLatestFormVersion	
	
			itm.UserProperties.Find("Division").Value = rs("division")
			itm.UserProperties.Find("ADDivision").Value = rs("division")
			
			itm.UserProperties.Find("NonPerson").Value = rs("NonPerson")
			itm.UserProperties.Find("ADNonPerson").Value = rs("NonPerson")
			
			itm.UserProperties.Find("PendingInactive").Value = rs("PendingInactive")
			itm.UserProperties.Find("ADPendingInactive").Value = rs("PendingInactive")
						
			itm.UserProperties.Find("SpeedDial").Value = rs("SpeedDial")			
			itm.UserProperties.Find("ADSpeedDial").Value = rs("SpeedDial")
			
			itm.Business2TelephoneNumber = rs("TelephoneExtn")
			itm.UserProperties.Find("ADTelephoneExtn").Value = rs("TelephoneExtn")	
			
			itm.UserProperties.Find("CreatedStatus").Value = rs("CreatedStatus")	
			
			If rs("SalesmenBrands") <> "" Then
				On Error Resume Next
				arSalesmenBrands = Split(rs("SalesmenBrands"), ",")
				For Each brand In arSalesmenBrands
					itm.UserProperties.Find("rep" & brand).Value = -1
				Next
				On Error Goto 0
			End If						
			
	itm.Save	
	
	If Err.Number <> 0 Then
		Log(intCounter & " - Error Creating Contact:" & rs("objectGUID") & "(" & rs("FullName") & ")" & " - " & Err.Number & " - " & Err.Description)
	End If
	
End Sub	

Function GetLatestFormVersion()
	On Error Resume Next
	Set itmnw = GetPublicFolder(strFolderPath).Items.Add("IPM.Contact.InternalGlobalContact")	
	itmnw.FileAs = "formcheck"
	itmnw.FullName = "formcheck"
	itmnw.UserProperties.Find("ADobjectGUID").Value = "formcheck"
	itmnw.Save	
	Set fldr = GetPublicFolder(strFolderPath)
	Set itms = fldr.Items
	Set itm = itms.Find("[ADobjectGUID] = " & Chr(34) & "formcheck" & Chr(34))
	Set objGeneralPage = itm.GetInspector.ModifiedFormPages("General")
	Set FormVersion = objGeneralPage.Controls("version")
	Log("GetLatestFormVersion:" & FormVersion.Caption)
	GetLatestFormVersion = FormVersion.Caption
	itm.Delete
	itm.Close(0)			
End Function

'Gets a Public folder based on a string path - e.g. 
'If Folder name in English is Public Folders\All Public Folders\Europeen Workflow then just paste in "European Workflow'
Public Function GetPublicFolder(strFolderPath)
   
    Dim colFolders 
    Dim objFolder 
    Dim arrFolders 
    Dim i 
    On Error Resume Next    
    
        Set objOL = CreateObject("Outlook.Application")
        objOL.Session.Logon GetDefaultOutlookProfile(), , False, True
        'WScript.Echo objOL.Session.CurrentUser.Name    
    
    strFolderPath = Replace(strFolderPath, "/", "\")
    arrFolders = Split(strFolderPath, "\")
    
    Set objFolder = objOL.Session.GetDefaultFolder(18)
    Set objFolder = objFolder.Folders.Item(arrFolders(0))
    If Not objFolder Is Nothing Then
        For i = 1 To UBound(arrFolders)
            Set colFolders = objFolder.Folders
            Set objFolder = Nothing
            Set objFolder = colFolders.Item(arrFolders(i))
            If objFolder Is Nothing Then
                Exit For
            End If
        Next
    End If
    Set GetPublicFolder = objFolder
    Set colFolders = Nothing
    Set objApp = Nothing
    Set objFolder = Nothing
    
	If Err.Number <> 0 Then
		Log(intCounter & "GetPublicFolder - Error - " & Err.Number & " - " & Err.Description)
	End If    
    
End Function

Function GetDefaultOutlookProfile()
	On Error Resume Next
	GetDefaultOutlookProfile = ""
	GetDefaultOutlookProfile = Trim(objWshShell.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\DefaultProfile"))
End Function			

Sub Log(strLog)
	Err.Clear
	On Error Resume Next	

	If objFSO.FolderExists("c:\Support") <> True Then	
		objFSO.CreateFolder "c:\Support"
	End If
	
	'Open / Create Text File
	Set fileLogon = objFSO.OpenTextFile("C:\Support\GroupStaff-PublicFolder-ContactLoader.txt", 8, True)

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

Sub WaitForOutlook()
  Err.Clear
  retval = objWshShell.Run (Chr(34) & "C:\Program Files (x86)\Microsoft Office\Office14\Outlook.exe" & Chr(34), 1, FALSE)
  Wscript.Sleep 3000
  On Error Resume Next	

	Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2")
 	   
    tfOutlookActive = FALSE
    intAddActivitesOutlookMod = 0
    Do While intAddActivitesOutlookMod < 100
       Set colItems = objWMIService.ExecQuery("Select * from Win32_Process where Description = 'OUTLOOK.EXE'",,48)       
       For Each objItem in colItems
       Return = objItem.GetOwner(strNameOfUser)
	   If UCase(objItem.Description) = "OUTLOOK.EXE" And UCase(strNameOfUser) = UCase(GetUserName()) Then
	      tfOutlookActive = TRUE
	      intAddActivitesOutlookMod = 100
	      Exit For
	   End If
       Next
       Wscript.Sleep 3000
       intAddActivitesOutlookMod = intAddActivitesOutlookMod + 1
    Loop
  
    If tfOutlookActive = TRUE Then
    	Log("    Found Outlook running as local user. Waiting 10 seconds...")
		WScript.Sleep 10000
		Log("        ... kicking off jobs.")
		intLatestFormVersion = CInt(GetLatestFormVersion())
		Call DeleteBogusContacts
		Call IterateSQLContacts
		Call DeleteNonRequiredSQLContacts		
		Exit Sub
	Else
		WScript.Quit		
    End If
  
  On Error Goto 0
End Sub	

Sub KillOutlook()

  Err.Clear
  On Error Resume Next	

	Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2")
 	   
    tfOutlookActive = FALSE
    intAddActivitesOutlookMod = 0
    Do While intAddActivitesOutlookMod < 100
       Set colItems = objWMIService.ExecQuery("Select * from Win32_Process where Description = 'OUTLOOK.EXE'",,48)       
       For Each objItem in colItems
       Return = objItem.GetOwner(strNameOfUser)
	   If UCase(objItem.Description) = "OUTLOOK.EXE" And UCase(strNameOfUser) = UCase(GetUserName()) Then
			Log("    Found Outlook running as local user. Killing...")
			objItem.Terminate
	      Exit Sub
	   End If
       Next
       intAddActivitesOutlookMod = intAddActivitesOutlookMod + 1
    Loop  
  On Error Goto 0

End Sub

Function GetUserName()
    On Error Resume Next
	Set objNet = Wscript.CreateObject("Wscript.Network")
    Do While strUserName = ""
		strUserName = objNet.UserName
		Counter = Counter + 1
		If Counter > 100000 Then
			WScript.Quit
		End If
    Loop
    GetUserName = strUserName
    On Error Goto 0
End Function