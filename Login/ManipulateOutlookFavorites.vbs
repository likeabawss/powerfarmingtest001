
	On Error Resume Next
	Call WaitForOutlook
	wscript.sleep(4000) '4 second
	
	Const olPublicFoldersAllPublicFolders = 18
	Const olFavoriteFoldersGroup = 4
		
	Dim olkApp, olkSes, olkFolder
	Set olkApp = CreateObject("Outlook.Application")
	Set olkSes = olkApp.GetNameSpace("MAPI")
	'Change the profile name on the next line'
	'olkSes.Logon , , False, True
	'olkSes.Logon "Outlook"
	'objOL.Session.Logon , , False, True	
	
	Set olkFolder = OpenOutlookFolder("\Public Folders - " & GetAddress() & "\All Public Folders\Group Staff")
	olkFolder.AddToPFFavorites
	Set olkFolder = OpenOutlookFolder("\Public Folders - " & GetAddress() & "\Favorites\Group Staff")
	AddFavoriteFolder olkFolder
	
	olkSes.Logoff
	Set olkApp = Nothing
	Set olkSes = Nothing
	Set olkFolder = Nothing
	
Sub WaitForOutlook()
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
	      tfOutlookActive = TRUE
	      intAddActivitesOutlookMod = 100
	      Exit For
	   End If
       Next
       Wscript.Sleep 3000
       intAddActivitesOutlookMod = intAddActivitesOutlookMod + 1
    Loop
  
    If tfOutlookActive = TRUE Then
		WScript.Sleep 10000
		Exit Sub
		'Call AddFavorites	
	Else
		WScript.Quit		
    End If
  
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
		
Sub AddFavoriteFolder(olkFolder)
    ' Purpose: Add a folder to Favorite Folders.'
    ' Written: 5/2/2009'
    ' Author:  BlueDevilFan'
    ' Outlook: 2007'
    Const olModuleMail = 0
    Const olFavoriteFoldersGroup = 4
        Dim olkPane, olkModule, olkGroup
    Set olkPane = olkApp.ActiveExplorer.NavigationPane
    Set olkModule = olkPane.Modules.GetNavigationModule(olModuleMail)
    Set olkGroup = olkModule.NavigationGroups.GetDefaultNavigationGroup(olFavoriteFoldersGroup)
    olkGroup.NavigationFolders.Add olkFolder
    Set olkPane = Nothing
    Set olkModule = Nothing
    Set olkGroup = Nothing
End Sub	

Function OpenOutlookFolder(strFolderPath)
    ' Purpose: Opens an Outlook folder from a folder path.'
    ' Written: 4/24/2009'
    ' Author:  BlueDevilFan'
    ' Outlook: All versions'
    Dim arrFolders, varFolder, bolBeyondRoot
    On Error Resume Next
    If strFolderPath = "" Then
        Set OpenOutlookFolder = Nothing
    Else
        Do While Left(strFolderPath, 1) = "\"
            strFolderPath = Right(strFolderPath, Len(strFolderPath) - 1)
        Loop
        arrFolders = Split(strFolderPath, "\")
        For Each varFolder In arrFolders
            Select Case bolBeyondRoot
                Case False
                    Set OpenOutlookFolder = olkSes.Folders(varFolder)
                    bolBeyondRoot = True
                Case True
                    Set OpenOutlookFolder = OpenOutlookFolder.Folders(varFolder)
            End Select
            If Err.Number <> 0 Then
                Set OpenOutlookFolder = Nothing
                Exit For
            End If
        Next
    End If
    On Error GoTo 0
End Function

'Function to get email address"
Function GetAddress()
	On Error Resume Next
	    Dim objSysInfo, objUser
	    Set objSysInfo = CreateObject("ADSystemInfo")
	    Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
	    GetAddress = objUser.EmailAddress 
End Function