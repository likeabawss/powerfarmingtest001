
'**********************************************************************************************
'
'				   Power Farming Group Software Deployment Script
'
'**********************************************************************************************

' Author: Michael Barrett
' Date: 13/05/2011

' Updates:
' ********
' 14/05/2011 - MB - Office 2010 Standard Laptop and Workstation Deployment

'Set scripting interaction
'Wscript.Interactive = FALSE

'Declares
Public objWshShell
Public objFSO
Public objNet
Public dtStartNow, dtLastNow

'Load Intrinsic Objects
Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objNet = Wscript.CreateObject("Wscript.Network")

    'Set Start Time
    dtStartNow = Now
    dtLastNow = Now

Call Main

Sub Main
	Log("  ")	
	Log("Software Deployment Script Starts")
	Log("Location: " & Location)
	Call MSOfficeStandard2010	
	Log("Software Deployment Script Ends")
	Log("  ")	
End Sub

Sub MSOfficeStandard2010

	'ChangeLog
	'**********
	'
	' 15/05/2010 - MEB - OS must be workstation check.
	' 15/05/2010 - MEB - Check Setup Location has Setup.exe
	' 02/06/2010 - MEB - Howard (SYDNEY) Installation Source Added
	' 23/11/2012 - KBS - Howard (Sydney) Installation Source moved off old server
	
	tfContinue = TRUE
	
	'Check is Workstation
	'If IsWorkstation = FALSE then 
	'	log("OS is NOT of type Workstation. Fail.")
	'	tfContinue = FALSE
	'Else
	'	log("OS is of type Workstation. Success.")	
	'End If
	
	'Set Location Specific Paths - Put the '\' at the end please!!	
	Select Case Location 
		Case "PFNZ"
			MSOfficeStandard2010_SourcePath = "\\192.168.50.230\publicstore\Office2010Standard\"
		Case "PFGAU_MAIN"
			MSOfficeStandard2010_SourcePath = "\\pfg-srv-005\installs$\Office2010Standard\"		
		Case "PFGAU_SERVICE"
			MSOfficeStandard2010_SourcePath = "\\pfgsrv-06\installs$\Office2010Standard\"
		Case "PFNZ_MABERS"
			MSOfficeStandard2010_SourcePath = ""
		Case "PFGAU_BRISBANE"
			MSOfficeStandard2010_SourcePath = ""
		Case "HOWARD_SYDNEY"
			MSOfficeStandard2010_SourcePath = "\\hau-srv-003\installs$\Office2010Standard\"		
		Case Else
			log("System was found to be in an invalid location. FAIL.")
			tfContinue = FALSE
	End Select
	
	'Check Presence of Executable
	If objFSO.FileExists(MSOfficeStandard2010_SourcePath & "Setup.exe") Then
		log("Setup.exe was found at " & MSOfficeStandard2010_SourcePath & ". Ok")
	Else
		log("Setup.exe was NOT found at " & MSOfficeStandard2010_SourcePath & ".")
		tfContinue = FALSE
	End If
	
	'Check Machine Naming Convention
	If InStr(GetComputerName, "PPFW") Or _
		InStr(GetComputerName, "NPFW") Or _
		 InStr(GetComputerName, "PPFG")	Or _
		  InStr(GetComputerName, "NPFG") Or _		 
			InStr(GetComputerName, "PHOW") Or _		 		  
			 InStr(GetComputerName, "NHOW") Or _
			  InStr(GetComputerName, "PMAB") Or _
			   InStr(GetComputerName, "NHENG") Or _
				 InStr(GetComputerName, "PHENG") Or _
					InStr(GetComputerName, "PFNZ-IT-") Or _
						InStr(GetComputerName, "PFNZ-SRV-") Then			   
		log("Passed Computer Naming Convention Check. Success.")
    Else
		log("Failed Computer Naming Convention Check. Fail.")
		tfContinue = FALSE
	End If

	'Check Installed Software. Currently we only want to upgrade Office Std. products to 2010.
	Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
	strComputer = "."
	strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
	strEntry1a = "DisplayName"
	strEntry1b = "QuietDisplayName"
	
	Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
	objReg.EnumKey HKLM, strKey, arrSubkeys
	For Each strSubkey In arrSubkeys
		intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
		If intRet1 <> 0 Then
			objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1
		End If
		If InStr(1, UCase(strValue1), UCase("Microsoft Office")) Then
			If InStr(1, UCase(strValue1), UCase("Update")) = 0 Then
				If InStr(1, UCase(strValue1), UCase("Standard")) Then
					If InStr(1, UCase(strValue1), UCase("2010")) Then
						tfOffice2010StdAlreadyInstalled = TRUE
						tfContinue = FALSE
					Else
						tfOffice2010StdReadyForUpgrade = TRUE
						tfOffice2010StdUnsupported = FALSE
						tfContinue = TRUE
					End If
				ElseIf InStr(1, UCase(strValue1), UCase("Professional")) Then
					tfOffice2010StdUnsupported = TRUE
					tfContinue = FALSE
				ElseIf InStr(1, UCase(strValue1), UCase("Premium")) Then
					tfOffice2010StdUnsupported = TRUE
					tfContinue = FALSE					
				End If
			End If
		Else		
			tfOffice2010StdReadyForFreshInstall = TRUE
		End If
	Next
	
	If tfContinue = TRUE and tfOffice2010StdUnsupported = FALSE Then
		'Put installation path and arguments here.
		If tfOffice2010StdReadyForUpgrade = TRUE then
			log("Ready for Office Upgrade. A version of Microsoft Office Standard was found on this machine.")
			retval = MsgBox("Would you like to upgrade to Office Standard 2010 now? " &_
				chr(10) & "If you choose No, you will be prompted to install at next logon.", 4,"Microsoft Office 2010 Upgrade.")
				MsgBox "SHUTDOWN ANY OFFICE APPLICATIONS YOU HAVE RUNNING NOW!!",0,"Microsoft Office 2010 Upgrade."
			Select Case retval
				Case 6
					'START UPGRADE				
					log(GetUserName & " chose to begin an upgrade now.")					
					objWshShell.Run MSOfficeStandard2010_SourcePath & "Setup.exe" &_
						" /adminfile " & MSOfficeStandard2010_SourcePath & "Updates\Office2010Standard.MSP", 7, TRUE
					log("Setup completed one way or another.")					
					exit sub
				case 7
					log(GetUserName & " cancelled the upgrade.")
					log("Setup cancelled.")					
					exit sub
			End Select						
		End If
		If tfOffice2010StdReadyForFreshInstall = TRUE then
			log("Ready for Office Installation. Microsoft Office was not found on this machine.")
			retval = MsgBox("Would you like to install Office Standard 2010 now? " &_
				chr(10) & "If you choose No, you will be prompted to install at next logon.", 4,"Microsoft Office 2010 Install.")			
			Select Case retval
				Case 6
					'START INSTALL
					log(GetUserName & " chose to begin installation now.")
					objWshShell.Run MSOfficeStandard2010_SourcePath & "Setup.exe" &_
						" /adminfile " & MSOfficeStandard2010_SourcePath & "Updates\Office2010Standard.MSP", 7, TRUE
					log("Setup completed one way or another.")					
					exit sub						
				case 7
					log(GetUserName & " cancelled the installation.")
					log("Setup cancelled.")					
					exit sub
			End Select
		End If
			
	Else
		If tfOffice2010StdAlreadyInstalled = TRUE Then
			log("Microsoft Office 2010 Std. was already found on this machine.")
		ElseIf tfOffice2010StdUnsupported = TRUE then
			log("Versions of Office other than Standard were found. These must be manually upgraded or reinstalled.")
		End If
		log("This machine failed one or more checks for auto-deployment Office Std. 2010. Please check one or more previously listed failure conditions. Quitting.")		
	End If

End Sub

Sub Log(strLog)
	Err.Clear
	On Error Resume Next	

	'Test for support folder, create if not found.
	If objFSO.FolderExists("C:\Support") = FALSE Then
		   objFSO.CreateFolder("C:\Support")
	End if
	
	'Open / Create Text File
	Set fileLogon = objFSO.OpenTextFile("C:\Support\SoftwareDeploymentLog.txt", 8, True)

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

Function GetIP

  'Clear Err Object
  Err.Clear

  'Enable Error Handling
  On Error Resume Next

  'Setup Constants
  Const wbemFlagReturnImmediately = &h10
  Const wbemFlagForwardOnly = &h20

  arrComputers = Array("127.0.0.1")
  For Each strComputer In arrComputers

     Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
     Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration", "WQL", _
					    wbemFlagReturnImmediately + wbemFlagForwardOnly)

     For Each objItem In colItems
     strTestIP = Join(objItem.IPAddress, ",")
     strIPEnabled = objItem.IPEnabled
     If strTestIP <> "0.0.0.0" And Trim(strTestIP) <> "" And strTestIP <> "255.255.255.255" And strIPEnabled = "True" Then
	strIPAddress = Join(objItem.IPAddress, ",")
	   If Instr(strIPAddress, "192.168.201.") <> 0 OR Instr(strIPAddress, "192.168.203.") <> 0 OR Instr(strIPAddress, "192.168.48.") <> 0 OR Instr(strIPAddress, "192.168.206.") <> 0 OR Instr(strIPAddress, "192.168.3.") <> 0 OR Instr(strIPAddress, "192.168.0.") <> 0 OR Instr(strIPAddress, "192.168.50.") <> 0 Then
		   GetIP = strIPAddress
	   End If
	End If
     Next
  Next

  'Disable Error Handling
  On Error Goto 0

End Function

Function Location
  Err.Clear
  On Error Resume Next

  'GetIP
  strIP = GetIP

  'Set Location
  If Instr(strIP, "192.168.48.") <> 0 OR Instr(strIP, "192.168.99.") <> 0 OR Instr(strIP, "192.168.50.") <> 0 Then
     'Set Location
     Location = "PFNZ"
  End If
  'Set Location
  If Instr(strIP, "192.168.203.") <> 0 Then
     'Set Location
     Location = "PFGAU_MAIN"
  End If
  'Set Location
  If Instr(strIP, "192.168.201.") <> 0 Then
     'Set Location
     Location = "PFGAU_SERVICE"
  End If
  'Set Location
  If Instr(strIP, "192.168.3.") <> 0 Then
     'Set Location
     Location = "PFNZ_MABERS"
  End If
  'Set Location
  If Instr(strIP, "192.168.206.") <> 0 Then
     'Set Location
     Location = "PFGAU_BRISBANE"
  End If
  'Set Location
  If Instr(strIP, "192.168.0.") <> 0 Then
     'Set Location
     Location = "HOWARD_SYDNEY"
  End If

  'Logging
  Log("Local IP Address: " & strIP)  
  On Error Goto 0
End Function

Function IsWorkstation
    Err.Clear
    On Error Resume Next

	   ProductType = UCase(objWshShell.RegRead(_
	      "HKLM\System\CurrentControlSet\Control\ProductOptions\ProductType"))
		   
    Select Case ProductType
      Case "SERVERNT"
		IsWorkstation = FALSE
      Case "WINNT"
		IsWorkstation = TRUE
      Case Else
		IsWorkstation = FALSE
    End Select	  
		
    On Error Goto 0
End Function

Function GetComputerName
    Err.Clear
    On Error Resume Next	
	GetComputerName = objWshShell.RegRead(_
		"HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")      	
    On Error Goto 0
End Function

Function GetUserName
    On Error Resume Next
    'This routine takes into account the fact that at logon, the UserName function is not
    'immediately available. Therefore, it loops (maximum 100000 times) until the variable has
    'been filled with some (any) text.
	
	Dim strUserName
	
    Do While strUserName = ""
       strUserName = objNet.UserName
      Counter = Counter + 1

      If Counter > 100000 Then
		strUserName = "Unknown"
		'Wscript.Quit
      End If
    Loop

	GetUserName = strUserName	
    On Error Goto 0
End Function