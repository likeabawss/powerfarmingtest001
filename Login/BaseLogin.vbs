
'**********************************************************************************************
'
'				   Power Farming Group Login Script
'
'**********************************************************************************************

' Author: Michael Barrett

'Set scripting interaction
'Wscript.Interactive = FALSE

'	ChangeLog
'	*********
'	16/05/2010 - MB - Changes To: Location_PFGAU_MAIN123
'					 	F: drive has been redirected to new server PFG-SRV-005
'				 	 	PFGSRV-01 has had all users removed from it's printer list by now. Removed removal entries.
'					 	PFG-SRV-005 is 64bit. We need to test how drivers work from 32bit clients before we can make
'					  		the new server the printer server for the site.
'					 	Removed the "net time" command. No OS's exist now that need it.
'					 	Removed Proxy Override. Proxy server is no longer used here.
'				     	Removed Power Policy and ShortDatePolicy. Imports as Reg entries fail using this method in Windows 7.
'					 	Removed IDSe42 branding and config setup automation. IDS is no longer actively in use.
'	16/05/2011 - MB - Removed Disclaimer Install Code
'	02/06/2011 - MB - Reconfigured HOWARD WSUS direction code. HOWSRV-01 is gone. HOWSRV-02 has taken over WSUS roles.
'	05/07/2011 - MB - MailArchive Outlook Shortcut has been rolled out to all NZ machines.
'	25/07/2011 - MB - Removed redundant SRV-12 Entries.
'	25/07/2011 - MB - Add PFW printer mappings through to designated PFNZ-SRV-029 ( 32bit print server )
'				      New procedure added to detect and log 64 / 32 OS architecture in use.
'	28/07/2010 - MB - Added CANO to Canterbury Retail
'	08/09/2011 - MB - Enabled Outlook Administration for Outlook 2007
'	15/09/2011 - MB - DisableInternetProxy lines were commented out. Noted that this was taking a lot of time to process 
'					  and is unneeded.
'	27/09/2011 - MB - Removed disabling of Windows Search.
'	28/09/2011 - MB - Removed Helpdesk Icon.
'	12/01/2011 - MB - New IP range in use at Howard Australia. 192.168.100.* is now local to Sydney.
'	07/11/2011 - CZ - Updated Howard_Sydney_Groupjobs to create a user folder on the L: (HAU-SRV-003\Data\home) if the user logs into HAU-SRV-004
'					  Updated the RedirectMyDoc function to include extra commands needed to redirect the 'Locations' shortcut in server 2008 
'					  for users logging in to HAU-SRV-004
'	15/12/2011 - MB - Added code to pickup the users current DefaultPrinter path-share before printer work begins, and set it back to the printer
'					  after printer work is complete in a bid to keep the users default printer static.
'	23/03/2014 - KS - Modified script to remove existing F: mapping for PFG from PFG-SRV-005\userdata and re-map to \\pfg-srv-017\pfg-srv-005
'	19/02/2015 - MB - Added proc EnableOutlookCachedExchangeModeAndPublicFolderFavoritesForLaptops to enable offline access to favoritised public folder.
'	13/07/2015 - MB - Added PFNZ-SRV-050 to Printer Installation.
'	20/07/2015 - MB - Removed Printer PFG-AUS-ADM from mapping.
'	02/02/2016 - KS - Added Auckland branch settings.
'   31/05/2016 - HT - Created elevated privledge for obgshell so that that regwrite will work for windows 8.1 and windows 10
'	15/06/2016 - KS - changed NZ Trend Officscan server to pfnz-srv-054 from pfnz-srv-028
'	06/07/2016 - KS - Changed Maber Motors F2 to use new v1.22 version
'	19/07/2016 - KS - Changed Auckland branch F2 to use new v1.22 version
'	25/07/2016 - KS - Changed Te Awamutu branch F2 to use new v1.22 version
'	08/08/2016 - KS - Changed Gisborne and Invercargill branches to F2 v1.22
'	09/08/2016 - KS - Changed Manawatu and Taranaki branches to F2 v1.22
'	31/08/2016 - KS - Added PF Wairarapa branch
'	07/09/2016 - KS - Changed Canterbury and West Coast branches to F2 v1.22
'	20/09/2016 - KS - Changed Gore and Otago branches to F2 v1.22
'	19/01/2017 - KS - Modified F2 to v1.22 in Sub GROUPJOB_RETAILADMIN
'	25/01/2017 - MB - Commented out all calls to WSUS_setup. Will be using GP soon instead.
'	03/02/2017 - HT - Removed PfgFultonDriveTS shortcuts for TPFG Clients and Setup to use PFGTS which is now pointing to pfgts.pfgaustralia.com.au not pfgsrv-01.powerfarming.co.nz
					 
'Dim obj vars
Public objWshShell		'WSH Shell
Public objFSO			'WSH FileSystemObject
Public objNET			'WSH Networking Object

'Dim vars
Public strLocal_OS		'The local operating system (server/wkstn)
Public strScriptPath		'The folder structure under this script
Public strSystemRoot		'The root directory of the local OS
Public strDebug 		'A temp var for writing data to a log file
Public strLocalTime		'Time in a standard format
Public strInstallApps		'List of installed applications
Public strUserName		'Network usernamePublic strComputerName 	'The computer NetBIOS name
Public strIEVersion		'Installed Internet Explorer Version Number
Public strRegSvr32Path
Public strServicePack
Public strCurrentDateTimeStamp
Public strComputerName
Public strLocation
Public strLogonServer
Public bitProcessor
Public tfTerminalUser
Public dtStartNow
Public dtLastNow
Public strDefaultPrinterPath
Dim bolRetAdm

'Drive Property Vars
Public strDriveLetter	       'Enumerating Disk Property
Public strDriveType	       'Enumerating Disk Property
Public strTotalDiskSpace       'Enumerating Disk Property
Public strFreeDiskSpace        'Enumerating Disk Property
Public strUsedDiskSpace        'Calculated Disk Property
Public strVolumeName	       'Volume Label
Public strUNCSharePath	       'For network drives, the UNC path for the driveletter

'Drive Property Arrays
Public arrayDriveLetter()	    'Enumerating Disk Property
Public arrayDriveType() 	    'Enumerating Disk Property
Public arrayDriveReady()
Public arrayTotalDiskSpace()	    'Enumerating Disk Property
Public arrayFreeDiskSpace()	    'Enumerating Disk Property
Public arrayUsedDiskSpace()
Public arrayVolumeName()
Public arrayUNCSharePath()

'Smtp Mailer Vars
Public strSMTPServer
Public strSMTPMailFrom
Public strSMTPSendTo
Public strSMTPMessageSubject
Public strSMTPMessageText

'**********************************************************************************************
'*											      *
'					    Call Main
'*											      *
'**********************************************************************************************

Sub Main

    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")

    'Set Start Time
    dtStartNow = Now
    dtLastNow = Now

    '
    '
    'Call Locals
    Call GetCurrentTimeStamp
    Call Local_OS	       'To ensure OS is Windows 98
    Call GetLocalSystemRoot    'To check install drive Free Space Req.
    Call GetUserName	       'Get the user name of the user logging in.
    Call GetComputerName	'Get the local NetBIOS computer name.
	
	'Stop and Exit for ThinStuff Boxes
	If strComputerName = "PFNZ-IT-004" Then
		Exit Sub	
	End If
	
	
    Call GetLogonServer
	Call GetOSBits
	Call RemoveServerOfficeScanRunRegEntries
	Call GetDefaultPrinterPath	


	 'Logging
	 Log("")
	 Log("*************************************************************************************************")
	 Log("")
	 Log("LogonScript Version: " & ThisScriptModifiedDateStamp())
	 Log("System " & UCASE(strComputerName) & " was logged into by " & UCase(strUserName))
	 Log("Operating System: " & strLocal_OS)
	 Log("Script Path: " & Wscript.ScriptFullName)
	 Log("Detected SystemRoot:" & strSystemRoot)
	 Log("Detected LogonServer:" & strLogonServer)
	 Log("OS Bit Architecture In Use:" & bitProcessor)
	 Log("Default Printer Path:" & strDefaultPrinterPath)	 

    Call GetServicePack
    Call GetRegSvr32RunPath
    Call GetLocation
    Call EnableOutlookCachedExchangeModeAndPublicFolderFavoritesForLaptops
	
    Call Set_Offline_Files_GoOfflineOnSlowLink

    On Error Resume Next

    '
    '
    'Special Groups
    Call SpecialGroups
	'Call ComputerSpecificJobs	

    'Cleanup Catchall
    objFSO.DeleteFile("c:\support\_res\*.log")
    objFSO.DeleteFile("c:\windows\system32\oga*.*")
    objFSO.DeleteFile("c:\windows\tasks\oga*.*")
    retval = objWshShell.Run ("C:\WINDOWS\$NtUninstallKB925877$\spuninst\spuninst.exe /quiet /passive /norestart", 0, FALSE)
	
	'Uninstall PowerShell 1.0
	Call CheckUninstallPowerShellV1
		
    'Run Location Specific Scripts
    Select Case strLocation
	   Case "PFG-AUSTRALIS"
		Call Location_PFGAU_AUSTRALIS
		Call CreateAXShortcuts
		Call OfficeScanUpdate
		'Call WSUS_Setup
		Call Cleanup
		'Call PFG_ADMIN_GroupJobs - Sets default printer based on dept. location. STOPPED
		'Call InstallSafeGuardProxyTool
		'Call CheckRepairSysAidAssetInventoryAgent
		'Call DisableServices
		'Call Enforce_Telnet_Access
		'Call OfficeScanUpdate
		'Call DesktopShortcutEnforcement 	
		'Call UpgradeInstallRes2
		'Call Cleanup	   
	   Case "PFH-RETAIL"
		'Call LogUserDistinguishedName	   
		Call PFH_Retail_2012
		Call SetPowerLinkHomePage
		Call CleanUp
	   Case "PFNZ"
		Call LogUserDistinguishedName
		Call Location_PFNZ
		'Call WSUS_Setup
		Call DisableServices
		Call Enforce_Telnet_Access
		Call CheckRepairSysAidAssetInventoryAgent
		Call OfficeScanUpdate
		Call DesktopShortcutEnforcement
		Call UpgradeInstallRes2
		'Call AddOutlookMailArchiveFolder
		Call SetDefaultPrinterFromLogoff
		Call CleanUp
	   Case "PFNZ_MABERS"
		Call Location_PFNZ_MABERS
		'Call WSUS_Setup
		Call DisableServices
		Call Enforce_Telnet_Access
		Call CheckRepairSysAidAssetInventoryAgent
		Call OfficeScanUpdate
		Call DesktopShortcutEnforcement
		Call UpgradeInstallRes2
		Call Cleanup
	   Case "PFGAU_MAIN"
		Call Location_PFGAU_MAIN
		Call PFG_ADMIN_GroupJobs
		Call InstallSafeGuardProxyTool
		Call CheckRepairSysAidAssetInventoryAgent
		'Call WSUS_Setup
		Call DisableServices
		Call Enforce_Telnet_Access
		Call OfficeScanUpdate
		Call DesktopShortcutEnforcement 	
		Call UpgradeInstallRes2
		Call Cleanup
	   Case "PFGAU_SERVICE"
	    Call CreateAXShortcuts
		Call Location_PFGAU_SERVICE
		Call PFG_PARTS_GroupJobs
		Call InstallSafeGuardProxyTool
		'Call WSUS_Setup
		Call DisableServices
		Call Enforce_Telnet_Access
		Call CheckRepairSysAidAssetInventoryAgent
		Call OfficeScanUpdate
		Call DesktopShortcutEnforcement
		Call UpgradeInstallRes2
		Call Cleanup
	   Case "PFGAU_BRISBANE"
		'Call WSUS_Setup
		Call CreateAXShortcuts
		Call DisableServices
		Call DesktopShortcutEnforcement
		Call UpgradeInstallRes2
		Call Cleanup
		Call Brisbane_GroupJobs
	   Case "HOWARD_SYDNEY"
		Call CreateAXShortcuts
		Call Howard_Sydney_GroupJobs
		'Call WSUS_Setup
		Call DisableServices
		Call Enforce_Telnet_Access
		Call CheckRepairSysAidAssetInventoryAgent
		'Call DemandSolutionsHAU
		Call DemandSolutionsPFGPFW
		Call OfficeScanUpdate
		Call DesktopShortcutEnforcement
		Call UpgradeInstallRes2
		Call Cleanup
	   Case Else
		Call Location_PFNZ
		Call OfficeScanUpdate
		Log("Unable to detect location. Default is PFNZ!")
		Call DisableServices
		Call CheckRepairSysAidAssetInventoryAgent
		Call Enforce_Telnet_Access
		Call DesktopShortcutEnforcement
		Call UpgradeInstallRes2
		Call Cleanup
    End Select
	
	'Call ManipulateOutlookFavorites
	
End Sub

Sub SetDefaultPrinterFromLogoff		

	On Error Goto 0
	Log("Entered SetDefaultPrinterFromLogoff")

    Do While strDefaultPrinterPath <> objWshShell.RegRead("HKCU\POWERFARMING\DefaultPrinterPath")
					
		If Trim(objWshShell.RegRead("HKCU\POWERFARMING\DefaultPrinterPath")) <> "" Then
			If InStr(objWshShell.RegRead("HKCU\POWERFARMING\DefaultPrinterPath"), "\\") > 0 Then
				objNET.SetDefaultPrinter objWshShell.RegRead("HKCU\POWERFARMING\DefaultPrinterPath")
			End If
		End If		

       'Setup UserName var
       If strDefaultPrinterPath <> objWshShell.RegRead("HKCU\POWERFARMING\DefaultPrinterPath") Then
      		Counter = Counter + 1 
      		Call GetDefaultPrinterPath      		
	   Else
	   		Exit Sub      		
       End If 
      
       'Shutdown script after more than 10000 loops
       If Counter > 15 Then
	 	Exit Sub
	   Else
	   	WScript.Sleep(1000) 	 	 
       End If

    Loop

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub ManipulateOutlookFavorites
	Err.Clear
	On Error Resume Next
	objWshShell.Run "Wscript.exe " & "\\powerfarming.co.nz\netlogon\svn-netlogon\login\ManipulateOutlookFavorites.vbs", 0, False
	Log("ManipulateOutlookFavorites Kicked Off.")
	On Error Goto 0
End Sub

Sub PFG_AX_Service_RemoteAppServer_Drive

	On Error Resume Next
	If strComputerName = "PFNZ-TS-003" Then
		objNet.MapNetworkDrive "X:", "\\pfnz-srv-028\PFG-AxCache-Service-Warranty"
	End If
	Log("PFG_AX_Service_RemoteAppServer_Drive: Mapped X")	

End Sub

Sub UnMapSDrive

	On Error Resume Next
	'Disconnect Wholesale M: drive if found.
	If objNet.FolderExists("S:") = True Then
	   Err.Clear
		Log("I: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "S:", TRUE, TRUE
		Do While objNet.FolderExists("S:")
			If objNet.FolderExists("S:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If
	
End Sub

Sub UnMapIDrive

	On Error Resume Next
	'Disconnect Wholesale I: drive if found.
	If objNet.FolderExists("I:") = True Then
	   Err.Clear
		Log("I: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "I:", TRUE, TRUE
		Do While objNet.FolderExists("I:")
			If objNet.FolderExists("I:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If
	
End Sub

Sub UnMapUDrive

	On Error Resume Next
	'Disconnect RetailData U: drive if found.
	If objNet.FolderExists("U:") = True Then
	   Err.Clear
		Log("U: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "U:", TRUE, TRUE
		Do While objNet.FolderExists("U:")
			If objNet.FolderExists("U:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If
	
End Sub

Sub PFH_Retail_2012
	On Error Resume Next
	
	'Detect	Login Server
	If strComputerName = "PFNZ-SRV-035" Or strComputerName = "PFNZ-SRV-038" Then	
		Log("User logged into 2012 Retail Terminal Environment.")	
	Else
		Exit Sub					
	End If	
	
	'Disconnect Wholesale U: drive if found.
	If objNet.FolderExists("U:") = True Then
	   Err.Clear
		Log("U: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "U:", TRUE, TRUE
		Do While objNet.FolderExists("U:")
			If objNet.FolderExists("U:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If	
	
	'Disconnect Wholesale I: drive if found.
	If objNet.FolderExists("I:") = True Then
	   Err.Clear
		Log("I: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "I:", TRUE, TRUE
		Do While objNet.FolderExists("I:")
			If objNet.FolderExists("I:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If
	bolRetAdm = False				
	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")
	For Each Group in User.Groups
		
		If Group.Name = "RetailAdmin" Then
			bolRetAdm = True
			Log("	 User member of: " & Group.Name)
		Else
		End if
	Next
	If bolRetAdm = True Then
			Call GROUPJOB_RETAILADMIN
	Else	
	For Each Group in User.Groups
		Log("	 User member of: " & Group.Name)
		Select Case Group.Name
			Case "stdgrp_MABERMOTORS"
				Call GROUPJOB_stdgrp_MABERMOTORS_2012
			Case "stdgrp_TeAwamutu"
				Call GROUPJOB_stdgrp_TEAWAMUTU_2012
			Case "stdgrp_AgEarth"
				Call GROUPJOB_stdgrp_NORTHLAND_2012
			'Case "stdgrp_AgriLife", "sec_Agrilife.Counter"
			'	Call GROUPJOB_stdgrp_AGRILIFE_2012
			Case "stdgrp_GISBORNE"
				Call GROUPJOB_stdgrp_GISBORNE_2012
			Case "stdgrp_PowerTrac"
				Call GROUPJOB_stdgrp_HAWKESBAY_2012
			Case "stdgrp_PFMANAWATU"
				Call GROUPJOB_stdgrp_MANAWATU_2012				
			Case "stdgrp_PREMIER"
				Call GROUPJOB_stdgrp_TARANAKI_2012					
			Case "stdgrp_BROWNWOODS", "sec_Timaru.Counter"
				Call GROUPJOB_stdgrp_TIMARU_2012
			Case "stdgrp_WestCoast"
				Call GROUPJOB_stdgrp_WESTCOAST_2012
			Case "stdgrpAshburton"
				Call GROUPJOB_stdgrp_ASHBURTON_2012				
			Case "stdgrp_CanterburyTractors"
				Call GROUPJOB_stdgrp_CANTERBURY_2012								
			Case "stdgrpOTAGO"
				Call GROUPJOB_stdgrp_OTAGO_2012								
			Case "stdgrp_PFAGSOUTHLAND","sec_Invercargill.Counter"
				Call GROUPJOB_stdgrp_INVERCARGILL_2012												
			Case "stdgrpPFGORE", "sec_Gore.Counter"
				Call GROUPJOB_stdgrp_GORE_2012
			Case "stdgrp_Auckland"
				Call GROUPJOB_stdgrp_AUCKLAND_2012
			Case "stdgrp_WAIRARAPA"
				Call GROUPJOB_stdgrp_WAIRARAPA_2012
			Case "stdgrp_HowardEngineering"
				Call GROUPJOB_stdgrp_HowardEngineering_2014
			Case "RetailSIPrinters"
				Log("       Installing All Retail Printers.")
				Call MapPrinter("\\PFNZ-SRV-028\HENGA")
				Call MapPrinter("\\PFNZ-SRV-028\HENGB")
				Call MapPrinter("\\PFNZ-SRV-028\HENGD")
				Call MapPrinter("\\PFNZ-SRV-028\HENGE")
				Call MapPrinter("\\PFNZ-SRV-028\HENGG")
				Call MapPrinter("\\PFNZ-SRV-028\HENGH")
				Call MapPrinter("\\PFNZ-SRV-028\HENGI")
				Call MapPrinter("\\PFNZ-SRV-028\RET-AKL-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-ASH-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-ASH-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-ASH-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-ASH-UBA")
				Call MapPrinter("\\PFNZ-SRV-028\RET-AWA-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-AWA-ADM2")
				Call MapPrinter("\\PFNZ-SRV-028\RET-AWA-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-AWA-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-BAL-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-CAN-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-CAN-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-CAN-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-CAN-SVC2")
				Call MapPrinter("\\PFNZ-SRV-028\RET-DAR-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-DAR-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-DAR-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-GIS-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-GIS-ADM2")
				Call MapPrinter("\\PFNZ-SRV-028\RET-GIS-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-GOR-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-GOR-PT2")
				Call MapPrinter("\\PFNZ-SRV-028\RET-GOR-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-HWK-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-HWK-COL")
				Call MapPrinter("\\PFNZ-SRV-028\RET-HWK-KON")
				Call MapPrinter("\\PFNZ-SRV-028\RET-HWK-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-INV-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-INV-ADM2")
				Call MapPrinter("\\PFNZ-SRV-028\RET-INV-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-INV-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-INV-WRK")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-ADM2")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-PRT2")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-SVC2")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-UGH")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MWT-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MWT-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-MWT-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-NTH-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-NTH-OFF")
				Call MapPrinter("\\PFNZ-SRV-028\RET-NTH-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-COL")
				Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-SAL")
				Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-URB")
				Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-ADM2")
				Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-SAL")
				Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-TIM-ADM")
				Call MapPrinter("\\PFNZ-SRV-028\RET-TIM-COP")
				Call MapPrinter("\\PFNZ-SRV-028\RET-TIM-PRT")
				Call MapPrinter("\\PFNZ-SRV-028\RET-TIM-SVC")
				Call MapPrinter("\\PFNZ-SRV-028\RET-WST-ADM")
				Log("       Completed done.")
		End Select
	Next
	End if
				
	Err.Clear	
End Sub



Sub NoAUAsDefaultShutdownOption
	On Error Resume Next	 
	Set objGroup = GetObject _
	  ("LDAP://cn=NoAUAsDefaultShutdownOption,dc=powerfarming,dc=co,dc=nz")
	objGroup.GetInfo	 
	arrMemberOf = objGroup.GetEx("member")
	For Each strMember in arrMemberOf
	  If InStr(strMember, strComputer) > 0 Then
		objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAUAsDefaultShutdownOption", 1, "REG_DWORD"	  
		Log("NoAUAsDefaultShutdownOption Enforced.")			
	  Else
		objWshShell.RegDelete "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAUAsDefaultShutdownOption"		
	  End If
	Next
End Sub

Sub SetPowerLinkHomePage
	On Error Resume Next
	objWshShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\Start Page", "http://powerlink.powerfarming.co.nz"
	Log("Start Page for IE set to: http://powerlink.powerfarming.co.nz")	
	On Error Goto 0
End Sub

Sub LogUserDistinguishedName
	'Dim strComputer, strUsername, objWMI, colUsers, objUser
	On Error Resume Next
	strComputer = "."
	Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\directory\LDAP")
	Set colUsers = objWMI.ExecQuery("SELECT * FROM ds_user where ds_sAMAccountName = '" & strUserName & "'")	   
	If colUsers.Count > 0 Then
		For Each objUser in colUsers	
		Log("DN:" & objUser.ds_distinguishedName)
		Next
	End If
End Sub

Sub CheckUninstallPowerShellV1

	On Error Resume Next
	If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\044DDFD2C33A4E243A66176EBC24640A\InstallProperties\DisplayName")) = "Windows PowerShell" Then
		retval = objWshShell.Run ("MsiExec.exe /X{2DFDD440-A33C-42E4-A366-71E6CB4246A0} /qn", 0, FALSE)	
	End If

End Sub

Sub GetDefaultPrinterPath
	On Error Resume Next
	Log("Entering GetDefaultPrinterPath")
	strComputer = "." 
	Set objWMIService = GetObject("winmgmts:" _ 
	   & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
	Set colInstalledPrinters =  objWMIService.ExecQuery _ 
	   ("Select * from Win32_Printer where Default = 'True'") 
	For Each objPrinter in colInstalledPrinters 
		strDefaultPrinterPath = objPrinter.ServerName & "\" & objPrinter.ShareName 		
		Exit Sub
	
	 Next 		 
End Sub

Sub CreateAXShortcuts

	On Error Resume Next
		
	'Copy Shortcuts to User Folder
	If objFSO.FolderExists(Ucase("C:\users\" & strusername)) Then
		usrpath = Ucase("c:\users\" & strusername)
	Else
		usrpath = UCase("c:\documents and settings\" & strusername)
	End If	

	Log("Creating AX Shortcuts." )	
	'Copy

	strDestinationPath=usrpath & "\RemoteAppRdpShortcuts"
	objFSO.DeleteFolder strDestinationPath
 
	If objFSO.FolderExists(strDestinationPath) THEN
		objFSO.DeleteFolder strDestinationPath, TRUE
		log("RemoteApps Folder Deleted")
	End If
	
	objFSO.CopyFolder "\\powerfarming.co.nz\netlogon\svn-netlogon\RemoteAppRdpShortcuts", strDestinationPath, True	
	
	strDesktop = objWshShell.SpecialFolders("Desktop")
	
	'Only Copy if Local AX Client Installation NOT Detected.
	If objFSO.FileExists("C:\Program Files (x86)\Microsoft Dynamics AX\50\Client\Bin\Ax32.exe") Or _
		objFSO.FileExists("C:\Program Files\Microsoft Dynamics AX\50\Client\Bin\Ax32.exe") Then
			objFSO.DeleteFile strDesktop & "\AX.lnk", TRUE	
			objFSO.DeleteFile strDesktop & "\AX-TEST-HAU.lnk", TRUE
			objFSO.DeleteFile strDesktop & "\AX Test.lnk", TRUE			
		Exit Sub
	End If	
	
	'Detect Location for Specific Shortcuts
	Select Case strLocation	
		Case "PFG-AUSTRALIS"
		
			'Create AX Shortcut
			If objFSO.FileExists(strDesktop & "\AX.lnk") THEN
				objFSO.DeleteFile strDesktop & "\AX.lnk", TRUE	
			End If
			Set objAX = objWshShell.CreateShortcut(strDesktop & "\AX Live.lnk")
			If objFSO.FileExists(strDesktop & "\AX Live.lnk") THEN
				objFSO.DeleteFile strDesktop & "\AX Live.lnk", TRUE	
				log ("Deleted AX Live shortcut")	
			End If	
			objAX.TargetPath = "c:\windows\system32\mstsc.exe"
			objAX.Arguments = chr(34) & usrpath & "\RemoteAppRdpShortcuts\FullColour\AX-Live-15bit.rdp" & chr(34)
			objAX.WindowStyle = 1
			objAX.IconLocation = usrpath & "\RemoteAppRdpShortcuts\DynamicsAX.ico"
			objAX.Description = "Dynamics AX Live"
			objAX.WorkingDirectory = ""
			objAX.Save	
			log ("Created AX Live 15bit shortcut")	

			'Create AX Test Shortcut
			Set objAX = objWshShell.CreateShortcut(strDesktop & "\AX Test.lnk")
			If objFSO.FileExists(strDesktop & "\AX Test.lnk") THEN
				objFSO.DeleteFile strDesktop & "\AX Test.lnk", TRUE	
				log ("Deleted AX Test shortcut")	
			End If
			objAX.TargetPath = "c:\windows\system32\mstsc.exe"
			objAX.Arguments = chr(34) & usrpath & "\RemoteAppRdpShortcuts\FullColour\AX-Test-15bit.rdp" & chr(34)
			objAX.WindowStyle = 1
			objAX.IconLocation = usrpath & "\RemoteAppRdpShortcuts\DynamicsAX.ico"
			objAX.Description = "Dynamics AX Test"
			bjAX.WorkingDirectory = ""
			objAX.Save	
			log ("Created AX Test 15bit shortcut")
		
		Case Else
	
			'Create AX Shortcut
			If objFSO.FileExists(strDesktop & "\AX.lnk") THEN
				objFSO.DeleteFile strDesktop & "\AX.lnk", TRUE	
			End If
			Set objAX = objWshShell.CreateShortcut(strDesktop & "\AX Live.lnk")
			If objFSO.FileExists(strDesktop & "\AX Live.lnk") THEN
				objFSO.DeleteFile strDesktop & "\AX Live.lnk", TRUE	
				log ("Deleted AX Live shortcut")	
			End If	
			objAX.TargetPath = "c:\windows\system32\mstsc.exe"
			objAX.Arguments = chr(34) & usrpath & "\RemoteAppRdpShortcuts\AX-Live.rdp" & chr(34)
			objAX.WindowStyle = 1
			objAX.IconLocation = usrpath & "\RemoteAppRdpShortcuts\DynamicsAX.ico"
			objAX.Description = "Dynamics AX Live"
			objAX.WorkingDirectory = ""
			objAX.Save	
			log ("Created AX Live 8bit shortcut")
			
			'Create AX Test Shortcut
			Set objAX = objWshShell.CreateShortcut(strDesktop & "\AX Test.lnk")
			If objFSO.FileExists(strDesktop & "\AX Test.lnk") THEN
				objFSO.DeleteFile strDesktop & "\AX Test.lnk", TRUE	
				log ("Deleted AX Test shortcut")	
			End If
			objAX.TargetPath = "c:\windows\system32\mstsc.exe"
			objAX.Arguments = chr(34) & usrpath & "\RemoteAppRdpShortcuts\AX-Test.rdp" & chr(34)
			objAX.WindowStyle = 1
			objAX.IconLocation = usrpath & "\RemoteAppRdpShortcuts\DynamicsAX.ico"
			objAX.Description = "Dynamics AX Test"
			bjAX.WorkingDirectory = ""
			objAX.Save	
			log ("Created AX Test 8bit shortcut")
			
	End Select
	
	If Ucase(strusername)="TSAUTOLOGON"  THEN	
		log("Username is autologon")

		Select Case Mid(strComputerName, 1, 4)
			Case "THAU"
				Set objAX = objWshShell.CreateShortcut(strDesktop & "\HowardTS.lnk")
				'Create HowardTS Desktop ShortCut
				If objFSO.FileExists(strDesktop & "\HowardTS.lnk") THEN
					objFSO.DeleteFile strDesktop & "\HowardTS.lnk", TRUE	
					log ("Deleted TSHoward shortcut")	
				End If
			
				'Create HowardTS Desktop Shortcut
				objAX.TargetPath = "c:\windows\system32\mstsc.exe"
				objAX.Arguments = chr(34) & usrpath & "\RemoteAppRdpShortcuts\HowardTS.rdp" & chr(34)
				objAX.WindowStyle = 1
				objAX.Description = "Howard Terminal Server"
				objAX.WorkingDirectory = ""
				objAX.Save	
				
				log ("Created HowardTS shortcut")	
				objWshShell.Run strDesktop & "\HowardTS.lnk", 3, FALSE	
			Case "TPFW"
				Set objAX = objWshShell.CreateShortcut(strDesktop & "\PfwTS.lnk")
				'Create HowardTS Desktop ShortCut
				If objFSO.FileExists(strDesktop & "\PfwTS.lnk") THEN
					objFSO.DeleteFile strDesktop & "\PfwTS.lnk", TRUE	
					log ("Deleted PfwTS shortcut")	
				End If
			
				'Create PfwTS Desktop Shortcut
				objAX.TargetPath = "c:\windows\system32\mstsc.exe"
				objAX.Arguments = chr(34) & usrpath & "\RemoteAppRdpShortcuts\PfwTS.rdp" & chr(34)
				objAX.WindowStyle = 1
				objAX.Description = "PowerFarming Terminal Server"
				objAX.WorkingDirectory = ""
				objAX.Save	
				
				log ("Created PfwTS shortcut")	
				objWshShell.Run strDesktop & "\PfwTS.lnk", 3, FALSE	

			Case "TPFG"
			
				Set objTS = objWshShell.CreateShortcut(strDesktop & "\PFGTS.lnk")
				'Delete Fulton drive shortcut
				If objFSO.FileExists(strDesktop & "\PfgFultonDriveTS.lnk") THEN
				objFSO.DeleteFile strDesktop & "\PfgFultonDriveTS.lnk", TRUE	
				log ("Deleted PfgFultonDriveTS shortcut")
				End If	
			
				'Create PFGTS Desktop Shortcut
				objTS.TargetPath = "c:\windows\system32\mstsc.exe"
				objTS.Arguments = chr(34) & usrpath & "\RemoteAppRdpShortcuts\PFGTS.rdp" & chr(34)
				objTS.WindowStyle = 1
				objTS.Description = "PFG Australia - Terminal Server"
				objTS.WorkingDirectory = ""
				objTS.Save	
				
				log ("Created PFGTS shortcut")	
				objWshShell.Run strDesktop & "\PFGTS.lnk", 3, FALSE					
			
		End Select	
		
	Else
	
		Select Case Mid(strComputerName, 1, 4)
			Case "NPFW"	
				Set objAX = objWshShell.CreateShortcut(strDesktop & "\PfwTS.lnk")
				'Create HowardTS Desktop ShortCut
				If objFSO.FileExists(strDesktop & "\PfwTS.lnk") THEN
					objFSO.DeleteFile strDesktop & "\PfwTS.lnk", TRUE	
					log ("Deleted PfwTS shortcut")	
				End If
			
				'Create PfwTS Desktop Shortcut
				objAX.TargetPath = "c:\windows\system32\mstsc.exe"
				objAX.Arguments = chr(34) & usrpath & "\RemoteAppRdpShortcuts\PfwTS.rdp" & chr(34)
				objAX.WindowStyle = 1
				objAX.Description = "PowerFarming Terminal Server"
				objAX.WorkingDirectory = ""
				objAX.Save	
			Case "PPFW"	
				Set objAX = objWshShell.CreateShortcut(strDesktop & "\PfwTS.lnk")
				'Create HowardTS Desktop ShortCut
				If objFSO.FileExists(strDesktop & "\PfwTS.lnk") THEN
					objFSO.DeleteFile strDesktop & "\PfwTS.lnk", TRUE	
					log ("Deleted PfwTS shortcut")	
				End If
			
				'Create PfwTS Desktop Shortcut
				objAX.TargetPath = "c:\windows\system32\mstsc.exe"
				objAX.Arguments = chr(34) & usrpath & "\RemoteAppRdpShortcuts\PfwTS.rdp" & chr(34)
				objAX.WindowStyle = 1
				objAX.Description = "PowerFarming Terminal Server"
				objAX.WorkingDirectory = ""
				objAX.Save							
		End Select
	
	End If	
		
End Sub

'Was ComputerSpecificJobs
Sub DemandSolutionsPFGPFW
	On Error Resume Next
	If UCASE(strComputerName) = "PFNZ-SRV-031" Or _
			UCASE(strComputerName) = "PFNZ-SRV-017" Or _ 
				InStr(UCASE(strComputerName),"NPFW") <> 0 Or _ 				
					UCASE(strComputerName) = "PFNZ-SRV-027" Then
		objNet.RemoveNetworkDrive "R:", True, True
		objNet.MapNetworkDrive "R:", "\\PFNZ-SRV-031\DSINTEGRATION" ,TRUE
		Log("PFG / PFW - Demand Solutions R: Drive Mapped.")
	End If
End Sub

'Was ComputerSpecificJobs
Sub DemandSolutionsHAU
	On Error Resume Next
	If UCASE(strComputerName) = "PFNZ-SRV-031" Or _
			UCASE(strComputerName) = "PFNZ-SRV-017" Or _ 
				InStr(UCASE(strComputerName),"NHAU") <> 0 Or _ 				
					UCASE(strComputerName) = "PFNZ-SRV-027" Then
					
				'Map FAMIS required drive
				If objNet.FolderExists("R:") = True Then
				   Err.Clear
					Log("	   R: drive found.. attempting to disconnect.")
					objNet.RemoveNetworkDrive "R:", TRUE, TRUE
				Do While objNet.FolderExists("R:")
					If objNet.FolderExists("R:") = False Then
					   Exit Do
					End If
					Wscript.Sleep 1000
					intCounter = intCounter + 1
					If intCounter = 10 Then
					   Exit Do
					End If
				Loop
				End If					
				
		'objNet.RemoveNetworkDrive "R:", True, True
		objNet.MapNetworkDrive "R:", "\\PFNZ-SRV-031\DSINTEGRATION-HOWARD" ,TRUE
		Log("HAU - Demand Solutions R: Drive Mapped.*")
	End If
	
	If strLocation = "HOWARD_SYDNEY" Then

				'Map FAMIS required drive
				If objNet.FolderExists("R:") = True Then
				   Err.Clear
					Log("	   R: drive found.. attempting to disconnect.")
					objNet.RemoveNetworkDrive "R:", TRUE, TRUE
				Do While objNet.FolderExists("R:")
					If objNet.FolderExists("R:") = False Then
					   Exit Do
					End If
					Wscript.Sleep 1000
					intCounter = intCounter + 1
					If intCounter = 10 Then
					   Exit Do
					End If
				Loop
				End If					
				
		'objNet.RemoveNetworkDrive "R:", True, True
		objNet.MapNetworkDrive "R:", "\\HAU-SRV-003\dsfiles" ,TRUE
		Log("HAU - Demand Solutions R: Drive Mapped.*")	
	
	End If
	
End Sub

Sub GetOSBits
  On Error Resume Next
  bitProcessor = trim(GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth)
  On Error Goto 0
End Sub

Sub SpecialGroups
  On Error Resume Next
  	objWshShell.RegWrite "HKCU\Software\Microsoft\Office\14.0\Outlook\Security\PublicFolderScript", 1, "REG_DWORD"
    If InStr(UCASE(strUserName), "KOORB") <> 0 Then
      'Enforce Internet Proxy Settings
      objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride", "192.*;168.8.152.101;txjsvsnwh001.tic.textron.com*;txtcor02.textronturf.com*;www.anz.com;216.148.248.132;202.2.59.40;deskbank1.westpac.co.nz;deskbank2.westpac.co.nz;daedong.co.kr;anz.co.nz;coddi.com;savsystem.merlo.com;161.71.70.*;cdserver2;srv-01;srv-02.powerfarming.co.nz;srv-04.powerfarming.co.nz;srv-05.powerfarming.co.nz;srv-06.powerfarming.co.nz;srv-07.powerfarming.co.nz;srv-08.powerfarming.co.nz;srv-09.powerfarming.co.nz;srv-10.powerfarming.co.nz;srv-11.powerfarming.co.nz;srv-06.powerfarming.co.nz;powerlink.powerfarming.co.nz;portal.powerfarming.co.nz;<local>"
      Call CleanUp
      Wscript.Quit
    End If
  On Error Goto 0
End Sub

Sub DesktopShortcutEnforcement
  On Error Resume Next
  'Delete Unwanted Shortcuts
  objFSO.DeleteFile(objWshShell.SpecialFolders("AllUsersDesktop") & "\SysAid.lnk")
  objFSO.DeleteFile(objWshShell.SpecialFolders("Desktop") & "\SysAid.lnk")
  objFSO.DeleteFile("C:\Users\Public\Public Desktop\SysAid.lnk")
  Log("Desktop Shorcuts Enforced.")
  On Error Goto 0
End Sub

Function ThisScriptModifiedDateStamp()
  On Error Resume Next
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFSO.GetFile(Wscript.ScriptFullName)
  ThisScriptModifiedDateStamp = CTimeStamp(CDate(objFile.DateLastModified))
  On Error Goto 0
End Function

Sub CheckRepairSysAidAssetInventoryAgent
  Err.Clear
  On Error Resume Next
  If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Ilient\Agent\MachineID")) = "00-F1-D0-00-F1-D0" Then
     objWshShell.RegWrite "HKLM\SOFTWARE\Ilient\Agent\MachineID", "", "REG_SZ"
     objWshShell.RegWrite "HKLM\SOFTWARE\Ilient\Agent\FirstTime", "Y", "REG_SZ"
     Log("SysAid Inventory Agent encountered an invalid MAC ID. ID cleared. Agent will attempt to redetect and correct.")
     retval = objWshShell.Run ("net stop sysaidagent", 0, TRUE)
     retval = objWshShell.Run ("net start sysaidagent", 0, TRUE)
  End If  
  'Enforce Other Settings
  objWshShell.RegWrite "HKLM\SOFTWARE\Ilient\Agent\AllowSubmitSR", "N", "REG_SZ"    
  On Error Goto 0
End Sub

Sub RemoveServerOfficeScanRunRegEntries
  Err.Clear
  On Error Resume Next
  'If Instr(UCase(strComputerName), "SRV") <> 0  Or _
	'	Instr(UCase(strComputerName), "TS") <> 0 Or _
	'		Instr(UCase(strComputerName), "IT") <> 0 Then
	'	objWshShell.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\OfficeScanNT Monitor"		
	'	Log("System Detected as Server: Removing OfficeScan RunReg Entry")  		
  'End If  
  On Error Goto 0
End Sub

Sub Set_Offline_Files_GoOfflineOnSlowLink
  Err.Clear
  On Error Resume Next
  objWshShell.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\NetCache\GoOfflineOnSlowLink"
  On Error Goto 0
End Sub

Sub Enforce_Telnet_Access
  On Error Resume Next
  Err.Clear
  If InStr(UCASE(strComputerName), "SRV") = 0 Then
    objWshShell.Run "sc config tlntsvr obj= LocalSystem", 7, True
    objWshShell.Run "sc config tlntsvr start= Auto", 7, True
    objWshShell.Run "tlntadmn config sec = -NTLM", 7, True
    objWshShell.Run "sc start tlntsvr", 7, True
      If objFSO.FileExists("C:\Windows\System32\curl.exe") = False Then _
	objFSO.CopyFile "\\powerfarming.co.nz\netlogon\commonapps\curl\curl.exe", "C:\Windows\System32\"
      If objFSO.FileExists("C:\Windows\System32\libssl32.dll") = False Then _
	objFSO.CopyFile "\\powerfarming.co.nz\netlogon\commonapps\curl\libssl32.dll", "C:\Windows\System32\"
      If objFSO.FileExists("C:\Windows\System32\libeay32.dll") = False Then _
	objFSO.CopyFile "\\powerfarming.co.nz\netlogon\commonapps\curl\libeay32.dll", "C:\Windows\System32\"
      If Err.Number = 0 Then
	  Log("Telnet Environment Setup Successfully")
      Else
	  Log("Error in Telnet Environment Setup")
      End If
    End If
    On Error Goto 0
End Sub

Sub Brisbane_GroupJobs
  Err.Clear
  On Error Resume Next
  On Error Goto 0
  
	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")
	
	'Loop through all groups user is a member of
	For Each Group in User.Groups
		'Logging
		Log("    User member of: " & Group.Name)

		'Do work based on group membership
		Select Case Group.Name
			Case "PFG_Marketing"
				Call GROUPJOB_PFG_MARKETING
		End Select
	Next
  'Clear Err Object
  Err.Clear

  'Enable Error Handling
  On Error Resume Next
    
  'ToDo:
  '*****
  '1. Full colour AX RDP shortcuts.   
  '
  '
  
	'Add WorkstationDR
'	If InStr(Ucase(strComputerName), "NPFG") <> 0 Or _
'		InStr(UCASE(strComputerName), "PPFG") <> 0 Then	
'		objWshShell.Run "Wscript.exe " & "\\powerfarming.co.nz\netlogon\svn-netlogon\login\PFG-WorkstationDR.vbs", 0, FALSE
'		Log("WorkstationDR Policy Enforced")
'	End If                                                    

	Select Case bitProcessor
		Case 32
			If Instr(1, UCase(strComputerName), "SRV") = 0 Or strComputerName = "PFGSRV-01" OR Instr(1, UCase(strUserName), "AUTO") = 0 Then
		
				Log("Starting 32bit PFG-BRISBANE Location Printer Configuration.")
								
				'Map 32bit Printers from PFG-SRV-022
			
				Log("Completed 32bit PFG-BRISBANE Location Printer Configuration.")
			End If
		Case 64
			If Instr(1, UCase(strComputerName), "PFG-SRV-005") = 0 Then
				Log("Starting 64bit PFG-BRISBANE Location Printer Configuration.")					
				
				'Map 64bit Printers from PFG-SRV-017
				
				Call MapPrinter("\\PFG-SRV-017\PFG-BRS-UMB")
												
				Log("Completed 64bit PFG-BRISBANE Location Printer Configuration.")
				
			End If
	End Select	
	
	'Delete old helpdesk icon.
	objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\HelpDesk.lnk")
	objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\HelpDesk.lnk")
	objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\IT Training Registration.lnk")
	objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\IT Training Registration.lnk")
		
	'Test for support folder, create if not found.
	If objFSO.FolderExists("C:\Support") = FALSE Then
		   objFSO.CreateFolder("C:\Support")
	End if

    '
    '
    'Enforce Outlook 2000 Settings
    'Outlook message arrival visual notification - ENABLED
    'objWshShell.RegWrite "HKCU\Software\Microsoft\Office\9.0\Outlook\Preferences\Notification", 1, "REG_DWORD"

    '
    '
    'Enable Outlook Administration through Exchange 2000 Server
    'Only implement if running Windows 2000
    'If strLocal_OS = "Windows 2000 Professional" And UCase(strComputerName) <> "SRV-09" And UCase(strComputerName) <> "SRV-11" Then
    '   objWshShell.Run "Regedit /S " & "\\PFGSRV-01\netlogon\EnableOutlookAdmin.reg", 7, TRUE
    'End If

  'Disable Error Handling
  On Error Goto 0

End Sub

Sub Howard_Sydney_GroupJobs

  Err.Clear
  On Error Resume Next
  
	'Add WorkstationDR
	If InStr(Ucase(strComputerName), "NHOW") <> 0 Or _
		InStr(UCASE(strComputerName), "PHOW") <> 0 Then	
		objWshShell.Run "Wscript.exe " & "\\powerfarming.co.nz\netlogon\svn-netlogon\login\HAU-WorkstationDR.vbs", 0, FALSE
		Log("WorkstationDR Policy Enforced")
	End If  
  
  Dim regids, regret, msiObject, dealercnctSettingsFile
  If ((InStr(UCASE(strComputerName), "SRV") = 0 AND InStr(UCASE(strComputerName), "HOW") or (strComputerName = "HAU-SRV-004"))) <> 0 Then

    'Enforce DealerConnect Connect Parameter
    dealercnctSettingsFile = "c:\Program Files\IDS Enterprise Systems Pty Ltd\Howard Australia Dealer Connect\settings.ini"
    If (objFSO.FileExists(dealercnctSettingsFile)) Then

      'Check Setting Exists
      Set dealercnctSettingsCheck = objFSO.OpenTextFile(dealerconctSettingsFile, 1)
      retval = dealercnctSettingsCheck.ReadAll
      dealercnctSettingsCheck.close

      If InString(1, UCase(retval), "SYSTEM=DEALER.HOWARD-AUSTRALIA.COM.AU") = 0 Then

	     'Rename existing file
	     objFSO.CopyFile dealercnctSettingsFile, "C:\settings.old"
	     'DeleteFile
	     objFSO.DeleteFile(dealercnctSettingsFile)
	     'Open renamed file for READING
	     Set txtIDSe42SettingsRD = objFSO.OpenTextFile("C:\settings.old", 1)
	     'Open new Settings.INI file for writing
	     Set txtIDSe42SettingsWR = objFSO.OpenTextFile(dealercnctSettingsFile, 8, TRUE)

	     'Loop through items in OLD file
	     Do While txtIDSe42SettingsRD.AtEndOfStream <> True
		'ReadLine
		strCurrentLine = txtIDSe42SettingsRD.ReadLine
		'Check for system entry
		If Instr(strCurrentLine, "SYSTEM=") <> 0 Then
		   'Change the value of strCurrentline to new DNS name.
		   strCurrentLine = "SYSTEM=dealer.howard-australia.com.au"
		    'Write line into new file
		    txtIDSe42SettingsWR.Write strCurrentLine
		    txtIDSe42SettingsWR.WriteLine
		Else
		    'Write line into new file
		    txtIDSe42SettingsWR.Write strCurrentLine
		    txtIDSe42SettingsWR.WriteLine
		End If
	     Loop
      End If
	  Log("System address for DealerConnect was updated to dealer.howard-australia.com.au")
    End If

    'Remove Unwanted Network Drives
    objNet.RemoveNetworkDrive "F:", True, True
    objNet.RemoveNetworkDrive "H:", True, True
    objNet.RemoveNetworkDrive "I:", True, True
    objNet.RemoveNetworkDrive "J:", True, True
    objNet.RemoveNetworkDrive "Y:", True, True
    objNet.RemoveNetworkDrive "L:", True, True

    'Map Main Network Drive
    'objNet.MapNetworkDrive "L:", "\\HAU-SRV-003\DATA" ,FALSE
	On Error Resume Next
	If objNet.FolderExists("L:") = True Then
	   Err.Clear
		Log("L: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "L:", TRUE, TRUE
		Do While objNet.FolderExists("L:")
			If objNet.FolderExists("L:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If			
	objNet.MapNetworkDrive "L:", "\\PFG-SRV-017\HAUDATA" ,FALSE

	'Map Printers
	If bitProcessor = 32 then
		
		'Remove Old Printers
		objNet.RemovePrinterConnection "\\HOWSRV-02\HOWADMIN1", TRUE, TRUE
		objNet.RemovePrinterConnection "\\HOWSRV-02\HOWADMIN2", TRUE, TRUE
		objNet.RemovePrinterConnection "\\HOWSRV-02\HOWADMIN3", TRUE, TRUE
		objNet.RemovePrinterConnection "\\HOWSRV-02\HOWCOLOUR", TRUE, TRUE
		objNet.RemovePrinterConnection "\\HOWSRV-02\HOWWHS1", TRUE, TRUE
		
		'Add New Printers
		Log("Starting 32bit HAU Location Printer Configuration.")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-AD1")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-AD2")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-AD3")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-AD4")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-CLR")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-DSP")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-ZB1")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-ZB2")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-ZB3")
		Call MapPrinter("\\HOWSRV-02\HAU-SYD-ZB4")
		Log("Completed 32bit HAU Location Printer Configuration.")
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-AD1")	
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-AD2")	
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-AD3")	
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-AD4")	
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-CLR")	
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-DSP")
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-ZB1")
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-ZB2")
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-ZB3")	
		'objNet.AddWindowsPrinterConnection("\\HOWSRV-02\HAU-SYD-ZB4")					
	ElseIf bitProcessor = 64 then
	
		'Connect New Printers
		Log("Starting 64bit HAU Location Printer Configuration.")		
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-AD1")
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-AD2")
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-AD3")
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-AD4")
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-CLR")
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-DSP")
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-ZB1")
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-ZB2")
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-ZB3")
		Call MapPrinter("\\HAU-SRV-003\HAU-SYD-ZB4")	
		Log("Completed 64bit HAU Location Printer Configuration.")		
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-AD1")	
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-AD2")	
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-AD3")	
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-AD4")	
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-CLR")	
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-DSP")	
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-ZB1")	
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-ZB2")	
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-ZB3")
		'objNet.AddWindowsPrinterConnection("\\HAU-SRV-003\HAU-SYD-ZB4")
	End If	

    Log("Howard Drives and Printers Mapped.")

    'Check for DealerConnect Install
    regids = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion" & _
	   "\Uninstall\{C1565A03-E284-43E2-92A7-8C067B83A53F}\DisplayName"
    regret = objWshShell.RegRead(regids)
    If trim(regret) <> "Howard Australia Dealer Connect" then
      Set msiObject = Wscript.CreateObject("WindowsInstaller.Installer")
      msiObject.UILevel = 3 + 64
      msiObject.InstallProduct "\\howsrv-02\netlogon\how\Howard_DealerConnect\Howard_DealerConnect_2007.msi"
      objWshShell.Run "\\howsrv-02\netlogon\how\Howard_DealerConnect\AutoRegister.exe", 7, FALSE
    End If
	
    'Delete Helpesk Webpage
    strDesktop = objWshShell.SpecialFolders("Desktop")
    objFSO.DeleteFile strDesktop & "\Helpdesk.lnk"

	'create a user directory in the home directory
	 temp=CreateHomeDir(strUsername, "\\HAU-SRV-003\Data\home\")
	 temp=CreateMediaDir(strUsername, "\\HAU-SRV-003\Data\home\")

		
  End If
  On Error Goto 0
End Sub

Function FamisInstalled()
  Err.Clear
  On Error Resume Next
  If objFSO.FileExists("C:\F2ProgramsLocal\FAMIS2000.exe") Then
     FamisInstalled = True
  Else
     FamisInstalled = False
  End If
  On Error Goto 0
End Function

Sub UpgradeInstallRes2

  Err.Clear
  On Error Resume Next
  'On Error Goto 0

  Log("Res2 Agent Upgrade Entered.")

  If strComputerName = "PFNZ-SRV-027" then
	Exit Sub  
  End If
  
  
  'Uninstall Res v1.0
  objWshShell.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\pfdaemon"
  objFSO.DeleteFolder("C:\Support\_res")
  objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\bginfo", "C:\Support\_res2\bginfo.wsf"

  objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\_res2"

  Exit Sub

  If InStr(UCASE(strComputerName), "SRV") = 0 Then



    If objFSO.FolderExists("C:\Support\_res2") = False Then
      objFSO.CreateFolder("C:\Support\_res2")
    End If
    If objFSO.FileExists("C:\Support\_res2\_daemon2.wsf") = False Then
	objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\_daemon2.wsf", "C:\Support\_res2\"
	objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\_res2", "C:\Support\_res2\_daemon2.wsf"
	Log("Res2 Agent Installed.")
    End If

    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\_jumpoff.wsf", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\fnLog.vbs", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\fnMEB.vbs", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\fnPfSched.vbs", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\fnPowerFarmingGroup.vbs", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\fnSysScript.vbs", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\ofscnroamchk.wsf", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\bginfo.exe", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\bginfo.wsf", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\pfg_bginfo.bgi", "C:\Support\_res2\"
    objFSO.CopyFile "\\powerfarming.co.nz\netlogon\_res2\pfnz_bginfo.bgi", "C:\Support\_res2\"

    Log("Res2 Agent Upgraded.")

    else
     Log("Res2 Agent Upgrade Skipped.")

  End If
  On Error Goto 0
End Sub

Sub Load_TNT_Regional_Settings
  Err.Clear
  On Error Resume Next
  objWshShell.RegWrite "HKCU\Control Panel\International\sShortDate", "dd/MM/yyyy", "REG_SZ"
  On Error Goto 0
End Sub

Sub WSUS_Setup

  'Clear Err Object
  Err.Clear

  'Enable Error Handling
  On Error Resume Next

  Log("WSUS Setup Started.")
  
  'Enforce Server Wsus Settings
  If Instr(UCase(strComputerName), "SRV") > 0 OR Instr(UCase(strComputerName), "TS") > 0 OR Instr(UCase(strComputerName), "DMZ") > 0 Then
  
	Log("WSUS server settings are being enforced.")
	
    'Run Location Specific Scripts
    Select Case strLocation
	   Case "PFNZ", "PFNZ_MABERS" 
	   
	     Log("WSUS Setup Started for pfnz pfnz_mabers.")
	   
		'Clear Err Object
		Err.Clear

		'Enable Error Handling
		On Error Resume Next

		'Get Configured WSUS Server
		strWSUSServer = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")
		strWSUSServer = Trim(strWSUSServer)
		If strWSUSServer <> "http://PFNZ-SRV-054.powerfarming.co.nz:8530" AND strWSUSServer <> "" Then

			'Re-Home Client
			If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFNZ-SRV-054.powerfarming.co.nz:8530" Then
			   'Log
			   Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Re-homing WSUS client " & UCase(strComputerName) & " to WSUS server PFNZ-SRV-054.powerfarming.co.nz.")
					   'Install Req. Keys
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 2, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
			   'Restart WSUS service
			   strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("wuauclt /detectnow", 0, TRUE)
			End If

		'First time requires special entries to be added.
		ElseIf strWSUSServer = "" Then

		  'Register Client
		  strKeyCheck = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\")
		  If Err.Number <> 0 Then
		     'Create Policy Keys
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\", 0, "REG_BINARY"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\", 0, "REG_BINARY"
		     'Log
		     Log("Created missing registry keys for WSUS.")
		  End If

		  'Clear Err Object
		  Err.Clear

		'Check current server
		If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFNZ-SRV-054.powerfarming.co.nz:8530" Then
		     'Log
		     Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Registering WSUS client " & UCase(strComputerName) & " to WSUS server http://PFNZ-SRV-054.powerfarming.co.nz:8530.")
					'Install Req. Keys
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 2, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
		     'Reregister with WSUS server
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AccountDomainSid"
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\PingID"
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\SusClientId"
		     'Restart WSUS service
		     strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
		     strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
		     strUpdate = objWshShell.Run ("wuauclt /resetauthorization /detectnow", 0, TRUE)
		  End If
		End If	   
	   Case Else
			Log("WSUS server settings are not configured for this site.")		
	End Select
  End If
  
  'Check System Type
  If Instr(UCase(strComputerName), "SRV") = False AND Instr(UCase(strComputerName), "TS") = False Then

	Log("WSUS client settings are being enforced here.")
  
    'Run Location Specific Scripts
    Select Case strLocation
	   Case "PFNZ", "PFNZ_MABERS"

		
		'Clear Err Object
		Err.Clear

		'Enable Error Handling
		On Error Resume Next

		'Get Configured WSUS Server
		strWSUSServer = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")
		strWSUSServer = Trim(strWSUSServer)
		If strWSUSServer <> "http://PFNZ-SRV-054.powerfarming.co.nz:8530" AND strWSUSServer <> "" Then
			Log("Check strWSUSServer. has value")
			'Re-Home Client
			If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFNZ-SRV-054.powerfarming.co.nz:8530" Then
			   'Log
			   Log("WSUS currently set to: " & WshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Re-homing WSUS client " & UCase(strComputerName) & " to WSUS server PFNZ-SRV-054.powerfarming.co.nz.")
					   'Install Req. Keys
				Set WshShell = WScript.CreateObject("WScript.Shell")
				If WScript.Arguments.Named.Exists("elevated") = False Then
					CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
					WScript.Quit
				Else
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
				   'Restart WSUS service
				   strUpdate = WshShell.Run ("net stop wuauserv", 0, TRUE)
				   strUpdate = WshShell.Run ("net start wuauserv", 0, TRUE)
				   strUpdate = WshShell.Run ("wuauclt /detectnow", 0, TRUE)
				End If
			End If
		'First time requires special entries to be added.

		Else
			if strWSUSServer = "" Then
				'Register Client
				strKeyCheck = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\")
				If Err.Number <> 0 Then
					'Create Policy Keys
					Set WshShell = WScript.CreateObject("WScript.Shell")
					If WScript.Arguments.Named.Exists("elevated") = False Then
						CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
						WScript.Quit
					Else
					WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\", 0, "REG_BINARY"
					WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\", 0, "REG_BINARY"
					End If
					'Log
					Log("Created missing registry keys for WSUS.")
				End If
				'Clear Err Object
				Err.Clear

				'Check current server
				If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFNZ-SRV-054.powerfarming.co.nz:8530" Then
					 'Log
					 Log("WSUS currently set to: " & WshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Registering WSUS client " & UCase(strComputerName) & " to WSUS server http://PFNZ-SRV-054.powerfarming.co.nz:8530.")
							'Install Req. Keys
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
					 'Reregister with WSUS server
					 WshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AccountDomainSid"
					 WshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\PingID"
					 WshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\SusClientId"
					 'Restart WSUS service
					 strUpdate = WshShell.Run ("net stop wuauserv", 0, TRUE)
					 strUpdate = WshShell.Run ("net start wuauserv", 0, TRUE)
					 strUpdate = WshShell.Run ("wuauclt /resetauthorization /detectnow", 0, TRUE)
				End If
			End If		
		End If

	   Case "PFNZ_MABERS"

		log("location set to pfnz_mabers")
		'Clear Err Object
		Err.Clear

		'Enable Error Handling
		On Error Resume Next

		'Get Configured WSUS Server
		strWSUSServer = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")
		strWSUSServer = Trim(strWSUSServer)
		If strWSUSServer <> "http://PFNZ-SRV-054.powerfarming.co.nz:8530" AND strWSUSServer <> "" Then

			'Re-Home Client
			If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFNZ-SRV-054.powerfarming.co.nz:8530" Then
			   'Log
			   Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Re-homing WSUS client " & UCase(strComputerName) & " to WSUS server PFNZ-SRV-028.powerfarming.co.nz.")
					   'Install Req. Keys
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
			   'Restart WSUS service
			   strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("wuauclt /detectnow", 0, TRUE)
			End If

		'First time requires special entries to be added.
		ElseIf strWSUSServer = "" Then

		  'Register Client
		  strKeyCheck = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\")
		  If Err.Number <> 0 Then
		     'Create Policy Keys
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\", 0, "REG_BINARY"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\", 0, "REG_BINARY"
		     'Log
		     Log("Created missing registry keys for WSUS.")
		  End If

		  'Clear Err Object
		  Err.Clear

		'Check current server
		If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFNZ-SRV-054.powerfarming.co.nz:8530" Then
		     'Log
		     Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Registering WSUS client " & UCase(strComputerName) & " to WSUS server http://PFNZ-SRV-054.powerfarming.co.nz:8530.")
					'Install Req. Keys
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFNZ-SRV-054.powerfarming.co.nz:8530", "REG_SZ"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
		     'Reregister with WSUS server
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AccountDomainSid"
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\PingID"
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\SusClientId"
		     'Restart WSUS service
		     strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
		     strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
		     strUpdate = objWshShell.Run ("wuauclt /resetauthorization /detectnow", 0, TRUE)
		  End If
		End If
		
	   Case "PFGAU_MAIN"

		'Clear Err Object
		Err.Clear

		'Enable Error Handling
		On Error Resume Next

		'Get Configured WSUS Server
		'strWSUSServer = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")
		'strWSUSServer = Trim(strWSUSServer)
		'If strWSUSServer <> "http://PFGSRV-05:8530" Then
			   'Log
		'	   Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Re-homing WSUS client " & UCase(strComputerName) & " to WSUS server PFGSRV-05.")
		'	   'Install Req. Keys
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFGSRV-05:80", "REG_SZ"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFGSRV-05:80", "REG_SZ"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
		'	   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
			   'Restart WSUS service
		'	   strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
		'	   strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
		'	   strUpdate = objWshShell.Run ("wuauclt /detectnow", 0, TRUE)
		'End If

	   Case "PFGAU_SERVICE"

		'Clear Err Object
		Err.Clear

		'Enable Error Handling
		On Error Resume Next

		'Get Configured WSUS Server
		strWSUSServer = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")
		strWSUSServer = Trim(strWSUSServer)
		If strWSUSServer <> "http://PFG-SRV-017.powerfarming.co.nz:8530" AND strWSUSServer <> "" Then

			'Re-Home Client
			If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFG-SRV-017.powerfarming.co.nz:8530" Then
			   'Log
			   Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Re-homing WSUS client " & UCase(strComputerName) & " to WSUS server PFG-SRV-017.powerfarming.co.nz.")
					   'Install Req. Keys
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFG-SRV-017.powerfarming.co.nz:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFG-SRV-017.powerfarming.co.nz:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
			   'Restart WSUS service
			   strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("wuauclt /detectnow", 0, TRUE)
			End If

		'First time requires special entries to be added.

		ElseIf strWSUSServer = "" Then

		  'Register Client
		  strKeyCheck = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\")
		  If Err.Number <> 0 Then
		     'Create Policy Keys
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\", 0, "REG_BINARY"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\", 0, "REG_BINARY"
		     'Log
		     Log("Created missing registry keys for WSUS.")
		  End If

		  'Clear Err Object
		  Err.Clear

		'Check current server
		If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFG-SRV-017.powerfarming.co.nz:8530" Then
		     'Log
		     Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Registering WSUS client " & UCase(strComputerName) & " to WSUS server http://PFG-SRV-017.powerfarming.co.nz:8530.")
					'Install Req. Keys
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFG-SRV-017.powerfarming.co.nz:8530", "REG_SZ"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFG-SRV-017.powerfarming.co.nz:8530", "REG_SZ"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
		     'Reregister with WSUS server
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AccountDomainSid"
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\PingID"
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\SusClientId"
		     'Restart WSUS service
		     strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
		     strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
		     strUpdate = objWshShell.Run ("wuauclt /resetauthorization /detectnow", 0, TRUE)
		  End If
		End If

	   Case "PFG-AUSTRALIS"
	   
		
		'Clear Err Object
		Err.Clear

		'Enable Error Handling
		On Error Resume Next

		'Get Configured WSUS Server
		strWSUSServer = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")
		strWSUSServer = Trim(strWSUSServer)
		If strWSUSServer <> "http://PFG-SRV-017.powerfarming.co.nz:8530" AND strWSUSServer <> "" Then
			Log("Check strWSUSServer. has value")
			'Re-Home Client
			If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFG-SRV-017.powerfarming.co.nz:8530" Then
			   'Log
			   Log("WSUS currently set to: " & WshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Re-homing WSUS client " & UCase(strComputerName) & " to WSUS server PFG-SRV-017.powerfarming.co.nz.")
					   'Install Req. Keys
				Set WshShell = WScript.CreateObject("WScript.Shell")
				If WScript.Arguments.Named.Exists("elevated") = False Then
					CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
					WScript.Quit
				Else
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFG-SRV-017.powerfarming.co.nz:8530", "REG_SZ"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFG-SRV-017.powerfarming.co.nz:8530", "REG_SZ"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
				   WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
				   'Restart WSUS service
				   strUpdate = WshShell.Run ("net stop wuauserv", 0, TRUE)
				   strUpdate = WshShell.Run ("net start wuauserv", 0, TRUE)
				   strUpdate = WshShell.Run ("wuauclt /detectnow", 0, TRUE)
				End If
			End If
		'First time requires special entries to be added.

		Else
			if strWSUSServer = "" Then
				'Register Client
				strKeyCheck = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\")
				If Err.Number <> 0 Then
					'Create Policy Keys
					Set WshShell = WScript.CreateObject("WScript.Shell")
					If WScript.Arguments.Named.Exists("elevated") = False Then
						CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
						WScript.Quit
					Else
					WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\", 0, "REG_BINARY"
					WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\", 0, "REG_BINARY"
					End If
					'Log
					Log("Created missing registry keys for WSUS.")
				End If
				'Clear Err Object
				Err.Clear

				'Check current server
				If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://PFG-SRV-017.powerfarming.co.nz:8530" Then
					 'Log
					 Log("WSUS currently set to: " & WshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Registering WSUS client " & UCase(strComputerName) & " to WSUS server http://PFG-SRV-017.powerfarming.co.nz:8530.")
							'Install Req. Keys
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://PFG-SRV-017.powerfarming.co.nz:8530", "REG_SZ"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://PFG-SRV-017.powerfarming.co.nz:8530", "REG_SZ"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
					 WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
					 'Reregister with WSUS server
					 WshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AccountDomainSid"
					 WshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\PingID"
					 WshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\SusClientId"
					 'Restart WSUS service
					 strUpdate = WshShell.Run ("net stop wuauserv", 0, TRUE)
					 strUpdate = WshShell.Run ("net start wuauserv", 0, TRUE)
					 strUpdate = WshShell.Run ("wuauclt /resetauthorization /detectnow", 0, TRUE)
				End If
			End If		
		End If
	   Case "PFGAU_BRISBANE"

		'Clear Err Object
		Err.Clear

		'Enable Error Handling
		On Error Resume Next

		'Get Configured WSUS Server
		strWSUSServer = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")
		strWSUSServer = Trim(strWSUSServer)
		If LCase(strWSUSServer) <> "http://pfgsrv-05:80" Then

			'Re-Home Client
			'If strWSUSServer = "http://PFGSRV-01:8530" OR strWSUSServer = "http://PFGSRV-02:8530" OR strWSUSServer = "http://SRV-01:8530" Then
			   'Log
			   Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Re-homing WSUS client " & UCase(strComputerName) & " to WSUS server PFGSRV-02.")
			   'Install Req. Keys
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://pfgsrv-05:80", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://pfgsrv-05:80", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
			   'Restart WSUS service
			   strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("wuauclt /detectnow", 0, TRUE)
			'End If

		Else

		  'Register Client
		  strKeyCheck = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\")
		  If Err.Number <> 0 Then
		     'Create Policy Keys
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\", 0, "REG_BINARY"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\", 0, "REG_BINARY"
		     'Log
		     Log("Created missing registry keys for WSUS.")
		  End If
		End If

		  'Clear Err Object
		  Err.Clear

	 Case "HOWARD_SYDNEY"

		'Clear Err Object
		Err.Clear

		'Enable Error Handling
		On Error Resume Next

		'Get Configured WSUS Server
		strWSUSServer = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")
		strWSUSServer = Trim(strWSUSServer)
		If strWSUSServer <> "http://HAU-SRV-003:8530" Then

			'Re-Home Client
			If strWSUSServer <> "http://HAU-SRV-003:8530" Then
			   'Log
			   Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Re-homing WSUS client " & UCase(strComputerName) & " to WSUS server HAU-SRV-003.")
			   'Install Req. Keys
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://HAU-SRV-003:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://HAU-SRV-003:8530", "REG_SZ"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
			   objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
			   'Restart WSUS service
			   strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
			   strUpdate = objWshShell.Run ("wuauclt /detectnow", 0, TRUE)
			End If

		  End If

		  'Register Client
		  strKeyCheck = objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\")
		  If Err.Number <> 0 Then
		     'Create Policy Keys
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\", 0, "REG_BINARY"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\", 0, "REG_BINARY"
		     'Log
		     Log("Created missing registry keys for WSUS.")
		  End If

		  'Clear Err Object
		  Err.Clear

		'Check current server
		If Trim(objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer")) <> "http://HAU-SRV-003:8530" Then
		     'Log
		     Log("WSUS currently set to: " & objWshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer") & " Registering WSUS client " & UCase(strComputerName) & " to WSUS server HAU-SRV-003.")
		   'Install Req. Keys
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer", "http://HAU-SRV-003:8530", "REG_SZ"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer", "http://HAU-SRV-003:8530", "REG_SZ"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AUOptions", 4, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AutoInstallMinorUpdates", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequency", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\DetectionFrequencyEnabled", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoRebootWithLoggedOnUsers", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\NoAutoUpdate", 0, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\RescheduleWaitTime", 10, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallDay", 0, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\ScheduledInstallTime", 12, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\SusServerVersion", 1, "REG_DWORD"
		     objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AutoUpdate\ConfigVer", 1, "REG_DWORD"
		     'Reregister with WSUS server
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\AccountDomainSid"
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\PingID"
		     objWshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\SusClientId"
		     'Restart WSUS service
		     strUpdate = objWshShell.Run ("net stop wuauserv", 0, TRUE)
		     strUpdate = objWshShell.Run ("net start wuauserv", 0, TRUE)
		     strUpdate = objWshShell.Run ("wuauclt /resetauthorization /detectnow", 0, TRUE)
		  End If

	 End Select
  End If
	Call NoAUAsDefaultShutdownOption
  'Disable Error Handling
  On Error Goto 0

End Sub

Sub GetLogonServer
  Err.Clear
  On Error Resume Next
  Set objWshShell = Wscript.CreateObject("Wscript.Shell")
  strLogonServer = objWshShell.ExpandEnvironmentStrings("%LogonServer%")

  'Problem with Detecting Logon Server At Howard.
  If strLocation = "HOWARD_SYDNEY" Then
     strLogonServer = "\\HOWSRV-02"
  End If

  On Error Goto 0
End Sub

Sub OfficeScanUpdate

	Err.Clear
	On Error Resume Next		
	
	'Notes:
	'Bitness dependant Trend Officescan entries:
	'	OfficeScan 10.0.1068 (SRV-14.powerfarming.co.nz) or OfficeScan 10.5.1083 (PFNZ-SRV-028.powerfarming.co.nz)
	'	**********************************************************************************************************
	'	32bit OS - HKLM\SOFTWARE\TrendMicro\PC-cillinNTCorp
	'	64bit OS - HKLM\SOFTWARE\Wow6432Node\TrendMicro\PC-cillinNTCorp
	
	Select Case bitProcessor
		Case 32
			osRegHome = "HKLM\SOFTWARE\TrendMicro\PC-cillinNTCorp"
		Case 64
			osRegHome = "HKLM\SOFTWARE\Wow6432Node\TrendMicro\PC-cillinNTCorp"		
	End Select
	
	If (Instr(UCase(strComputerName), "SRV") = False And Instr(UCase(strComputerName), "UPFW005") = False) Then

		Log("OfficeScan Proc")
	
		'OfficeScan - Power Farming ( Wholesale )
		If strLocation = "PFNZ" OR strLocation = "PFNZ_MABERS" Then
			If InStr(UCASE(strComputerName), "NPFW") <> 0 Or _
				InStr(UCASE(strComputerName), "PPFW") <> 0 Or _
				InStr(UCASE(strComputerName), "PMAB") <> 0 Or _	   
				InStr(UCASE(strComputerName), "UPFW") <> 0 Or _	   
				InStr(UCASE(strComputerName), "NMAB") <> 0 Then			

'				If objWshShell.RegRead(osRegHome & "\CurrentVersion\InstDate") < 20120116 Then
'					Log("Forcing Reinstall. Running OfficeScan Removal/Install/Update from SRV-14 to PFNZ-SRV-028.")
'					retval = objWshShell.Run ("net stop ntrtscan", 0, TRUE)
'					retval = objWshShell.Run ("net stop TmProxy", 0, TRUE)
'					retval = objWshShell.Run ("net stop tmlisten", 0, TRUE)
'					objWshShell.Run "\\powerfarming.co.nz\netlogon\CmnUnins.exe", 7, TRUE
'					objWshShell.Run "\\pfnz-srv-054\ofcscan\AUTOPCC.EXE", 7, FALSE
'				ElseIf Trim(objWshShell.RegRead(osRegHome & "\CurrentVersion\Misc.\ProgramVer")) = "" Then
'					Log("Running OfficeScan Removal/Install/Update from PFNZ-SRV-054.")		
'					objWshShell.Run "\\pfnz-srv-054\ofcscan\AUTOPCC.EXE", 7, FALSE	   
'				End If
			End If
		End If

		'OfficeScan - PFG Australia - Australis Ave
		If strLocation = "PFG-AUSTRALIS" Then
			If InStr(UCASE(strComputerName), "NPFG") <> 0 Or _
				InStr(UCASE(strComputerName), "PPFG") <> 0 Or _
					InStr(UCASE(strComputerName), "NHOW") <> 0 Or _
						InStr(UCASE(strComputerName), "PHOW") <> 0 Then				
							If objWshShell.RegRead(osRegHome & "\CurrentVersion\InstDate") < 20120116 Then	  
								Log("Forcing Reinstall. Running OfficeScan Removal/Install/Update from PFG-SRV-017.")
								retval = objWshShell.Run ("net stop ntrtscan", 0, TRUE)
								retval = objWshShell.Run ("net stop TmProxy", 0, TRUE)
								retval = objWshShell.Run ("net stop tmlisten", 0, TRUE)	     
								objWshShell.Run "\\powerfarming.co.nz\netlogon\CmnUnins.exe", 7, TRUE
								objWshShell.Run "\\PFG-SRV-017\ofcscan\AUTOPCC.EXE", 7, FALSE
							ElseIf Trim(objWshShell.RegRead(osRegHome & "\CurrentVersion\Misc.\ProgramVer")) = "" Then
								Log("Installing OfficeScan from PFG-SRV-017.")	       
								objWshShell.Run "\\pfg-srv-017\ofcscan\AUTOPCC.EXE", 7, FALSE		 
							End If	    
			End If	    
		End If
		
		'OfficeScan - PFG Australia
		If strLocation = "PFGAU_MAIN" OR strLocation = "PFGAU_SERVICE" Then
			If InStr(UCASE(strComputerName), "NPFG") <> 0 Or _
				InStr(UCASE(strComputerName), "PPFG") <> 0 Then
				If objWshShell.RegRead(osRegHome & "\CurrentVersion\InstDate") < 20120116 Then	  
					Log("Forcing Reinstall. Running OfficeScan Removal/Install/Update from PFG-SRV-017.")
					retval = objWshShell.Run ("net stop ntrtscan", 0, TRUE)
					retval = objWshShell.Run ("net stop TmProxy", 0, TRUE)
					retval = objWshShell.Run ("net stop tmlisten", 0, TRUE)	     
					objWshShell.Run "\\powerfarming.co.nz\netlogon\CmnUnins.exe", 7, TRUE
					objWshShell.Run "\\PFG-SRV-017\ofcscan\AUTOPCC.EXE", 7, FALSE
				ElseIf Trim(objWshShell.RegRead(osRegHome & "\CurrentVersion\Misc.\ProgramVer")) = "" Then
					Log("Installing OfficeScan from PFG-SRV-017.")	       
					objWshShell.Run "\\PFG-SRV-017\ofcscan\AUTOPCC.EXE", 7, FALSE		 
				End If	    
			End If	    
		End If		   

		'OfficeScan - Howard Australia
		If strLocation = "HOWARD_SYDNEY" Then
			If InStr(UCASE(strComputerName), "NHOW") <> 0 Or _
				InStr(UCASE(strComputerName), "PHOW") <> 0 Then
				If objWshShell.RegRead(osRegHome & "\CurrentVersion\InstDate") < 20120116 Then	  
					Log("Forcing Reinstall. Running OfficeScan Removal/Install/Update from HAU-SRV-003.")
					retval = objWshShell.Run ("net stop ntrtscan", 0, TRUE)
					retval = objWshShell.Run ("net stop TmProxy", 0, TRUE)
					retval = objWshShell.Run ("net stop tmlisten", 0, TRUE)	     
					objWshShell.Run "\\powerfarming.co.nz\netlogon\CmnUnins.exe", 7, TRUE
					objWshShell.Run "\\hau-srv-003\ofcscan\AUTOPCC.EXE", 7, FALSE
				ElseIf Trim(objWshShell.RegRead(osRegHome & "\CurrentVersion\Misc.\ProgramVer")) = "" Then
					Log("Installing OfficeScan from HAU-SRV-003.")	       
					objWshShell.Run "\\hau-srv-003\ofcscan\AUTOPCC.EXE", 7, FALSE		 
				End If	    
			End If	    
		End If		   

       
	End If

  On Error Goto 0
End Sub

Sub Location_PFNZ 

  'Clear Err Object
  Err.Clear

  'Enable Error Handling
  On Error Resume Next
    
  'Map Network Drives
  objNet.RemoveNetworkDrive "I:", True, True
  'objNet.MapNetworkDrive "I:", "\\SRV-01\VOL1" ,TRUE   
  objNet.MapNetworkDrive "I:", "\\PFNZ-SRV-028\PFWDATA" ,TRUE 

	'
	'
	'Enable Outlook Administration for Outlook 2007	
	objWshShell.RegWrite "HKCU\Software\Policies\Microsoft\Office\12.0\Outlook\Security\AdminSecurityMode", 1, "REG_DWORD"	
  
	'Add WorkstationDR
	If InStr(Ucase(strComputerName), "NPFW") <> 0 Or _
		InStr(UCASE(strComputerName), "UPFW") <> 0 Or _
		InStr(UCASE(strComputerName), "PPFW") <> 0 Then	
		objWshShell.Run "Wscript.exe " & "\\powerfarming.co.nz\netlogon\svn-netlogon\login\PFW-WorkstationDR.vbs", 0, FALSE
		Log("WorkstationDR Policy Enforced")
	End If
  
	'Add Printers
	If InStr(UCASE(strComputerName), "NPFW") <> 0 Or _
		 InStr(UCASE(strComputerName), "UPFW") <> 0 Or _
	     InStr(UCASE(strComputerName), "PPFW") <> 0 Or _
		 InStr(UCASE(strComputerName), "PFNZ-SRV-017") <> 0 Or _		 
		 InStr(UCASE(strComputerName), "PFNZ-SRV-003") <> 0 Or _		 
		 InStr(UCASE(strComputerName), "PFNZ-SRV-050") <> 0 Or _		 
		 InStr(UCASE(strComputerName), "PFNZ-SRV-027") <> 0 Then

			'Map Printers
			If bitProcessor = 32 then				
				Log("Starting 32bit PFNZ Location Printer Configuration.")
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-ACC")				
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-ASS")
				'Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-COP")
				objNet.RemovePrinterConnection "\\PFNZ-SRV-028\PFW-MVL-COP", TRUE, TRUE	
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-COP2")
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-COP2-BW")
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-DES")
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-LOG")
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-LOG1")
				'Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-PRT")
				objNet.RemovePrinterConnection "\\PFNZ-SRV-029\PFW-MVL-PRT", TRUE, TRUE
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-REC")
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-SHP")
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-WAR")
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-LP2")			
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-COL2")
				objNet.RemovePrinterConnection "\\PFNZ-SRV-029\PFW-MVL-COL", TRUE, TRUE
				Call MapPrinter("\\PFNZ-SRV-029\PFW-MVL-PB1")
				Call MapPrinter("\\PFNZ-SRV-029\PFW-CHC-WHS")
				Log("Completed 32bit PFNZ Location Printer Configuration.")
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-ACC")
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-ASS")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-COP")  	
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-DES")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-LOG")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-LOG1")  				
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-PRT")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-REC")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-SHP")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-WAR")     
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-LP2")				
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-COL")
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-PB1")
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-CHC-WHS")
			ElseIf bitProcessor = 64 then
				Log("Starting 64bit PFNZ Location Printer Configuration.")	
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-ACC")				
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-ASS")
				'Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-COP")
				objNet.RemovePrinterConnection "\\PFNZ-SRV-028\PFW-MVL-COP", TRUE, TRUE	
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-COP2")
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-COP2-BW")
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-DES")
				objNet.RemovePrinterConnection "\\PFNZ-SRV-028\PFW-MVL-LOG", TRUE, TRUE
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-LOG1")				
				'Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-PRT")
				objNet.RemovePrinterConnection "\\PFNZ-SRV-028\PFW-MVL-PRT", TRUE, TRUE
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-REC")
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-SHP")
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-WAR")
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-LP2")
				objNet.RemovePrinterConnection "\\PFNZ-SRV-028\PFW-MVL-COL", TRUE, TRUE		
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-COL2")
				Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-PB1")
				Call MapPrinter("\\PFNZ-SRV-028\PFW-CHC-WHS")
				Call MapPrinter("\\PFNZ-SRV-028\PFW-CHC-LP1")
				Log("Completed 64bit PFNZ Location Printer Configuration.")					
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-ACC")
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-ASS")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-COP")  	
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-DES")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-LOG")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-LOG1")  				
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-PRT")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-REC")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-SHP")  
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-WAR")   
 				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-LP2")				
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-COL")
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-MVL-PB1")
				'objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-028\PFW-CHC-WHS")				
			End If
			
	End If		
			
    Call DisableInternetProxy
	 'Log
	 Log("Completed DisableInternetProxy")
    Call InstallSafeGuardProxyTool
	 'Log
	 Log("Completed InstallSafeGuardProxyTool")
    Call KillAdWatch
	 'Log
	 Log("Completed KillAdWatch")
    Call DisableServices
	 'Log
	 Log("Completed DisableServices")
    Call SetupHelpDesk
	 'Log
	 Log("Completed SetupHelpDesk")
    Call GroupJobs
	 'Log
	 Log("Completed GroupJobs")
    Call EnforceSettings
	 'Log
	 Log("Completed EnforceSettings")

  'Disable Error Handling
  On Error Goto 0

End Sub

Sub Location_PFGAU_AUSTRALIS

	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")
	
	'Loop through all groups user is a member of
	For Each Group in User.Groups
		'Logging
		Log("    User member of: " & Group.Name)

		'Do work based on group membership
		Select Case Group.Name
			Case "PFG_Marketing"
				Call GROUPJOB_PFG_MARKETING
		End Select
	Next
  'Clear Err Object
  Err.Clear

  'Enable Error Handling
  On Error Resume Next
    
  'ToDo:
  '*****
  '1. Full colour AX RDP shortcuts.   
  '
  '
  
	'Add WorkstationDR
	If InStr(Ucase(strComputerName), "NPFG") <> 0 Or _
		InStr(UCASE(strComputerName), "PPFG") <> 0 Then	
		objWshShell.Run "Wscript.exe " & "\\powerfarming.co.nz\netlogon\svn-netlogon\login\PFG-WorkstationDR.vbs", 0, FALSE
		Log("WorkstationDR Policy Enforced")
	End If  

    'Map drives
   ' objNet.RemoveNetworkDrive "F:", TRUE, TRUE
   ' objNet.MapNetworkDrive "F:", "\\PFG-SRV-005\UserData" ,TRUE

    objNet.RemoveNetworkDrive "F:", TRUE, TRUE
  	On Error Resume Next
  	If objNet.FolderExists("F:") = True Then
	   Err.Clear
		Log("F: drive found.. attempting to disconnect.")
                objNet.RemoveNetworkDrive "F:", TRUE, TRUE
                Do While objNet.FolderExists("F:")
                        If objNet.FolderExists("F:") = False Then
                           Exit Do
                        End If
                        Wscript.Sleep 1000
                        intCounter = intCounter + 1
                        If intCounter = 10 Then
                           Exit Do
                        End If
                Loop      
        End If                                                    
    objNet.MapNetworkDrive "F:", "\\PFG-SRV-017\PFG-SRV-005", TRUE

	
	'Map drives
	'objNet.MapNetworkDrive "G:", "\\PFG-SRV-017\PFG-FUL-DTA", TRUE
    objNet.RemoveNetworkDrive "G:", TRUE, TRUE
	On Error Resume Next
	If objNet.FolderExists("G:") = True Then
	   Err.Clear
		Log("G: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "G:", TRUE, TRUE
		Do While objNet.FolderExists("G:")
			If objNet.FolderExists("G:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If				
    objNet.MapNetworkDrive "G:", "\\PFG-SRV-017\PFG-SRV-011", TRUE
	
	'Map drives
    'objNet.RemoveNetworkDrive "L:", TRUE, TRUE
    'objNet.MapNetworkDrive "L:", "\\HAU-SRV-003\DATA", TRUE	
	On Error Resume Next
	If objNet.FolderExists("L:") = True Then
	   Err.Clear
		Log("L: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "L:", TRUE, TRUE
		Do While objNet.FolderExists("L:")
			If objNet.FolderExists("L:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If
	
	Select Case bitProcessor
		Case 32
			If Instr(1, UCase(strComputerName), "SRV") = 0 Or strComputerName = "PFGSRV-01" OR Instr(1, UCase(strUserName), "AUTO") = 0 Then
		
				Log("Starting 32bit PFG-AUSTRALIS Location Printer Configuration.")
								
				'Map 32bit Printers from PFG-SRV-022
				'Call MapPrinter("\\PFG-SRV-022\PFG-AUS-ADM")
				objNet.RemovePrinterConnection "\\PFG-SRV-022\PFG-AUS-ADM", TRUE, TRUE
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-ASS")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-LOG")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-APC")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-MGT")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-MKT")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-SHP")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-OFF")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-UPC")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-DSP")
				'Call MapPrinter("\\PFG-SRV-022\PFG-FUL-WS1")
				'Call MapPrinter("\\PFG-SRV-022\PFG-FUL-WS2")				
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-PRT")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-DSP")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-REC")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-DPC")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-WAR")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-WS1")
				Call MapPrinter("\\PFG-SRV-022\PFG-AUS-WS2")
			
				Log("Completed 32bit PFG-AUSTRALIS Location Printer Configuration.")
			End If
		Case 64
			If Instr(1, UCase(strComputerName), "PFG-SRV-005") = 0 Then
				Log("Starting 64bit PFG-AUSTRALIS Location Printer Configuration.")							
				
				'Remove problem printers
				objNet.RemovePrinterConnection "\\PFG-SRV-017\PFG-AUS-ADM", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFG-SRV-017\PFG-AUS-LOG", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFG-SRV-017\PFG-AUS-OFF", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFG-SRV-017\PFG-AUS-WAR", TRUE, TRUE
				
				
				'Map 64bit Printers from PFG-SRV-017
				
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-ASS")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-LOG")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-APC")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-MGT")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-MKT")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-SHP")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-OFF")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-UPC")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-DSP")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-DSP2")				
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-PRT")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-DSP")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-REC")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-DPC")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-WAR")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-WS1")
				Call MapPrinter("\\PFG-SRV-017\PFG-AUS-WS2")
				Call MapPrinter("\\PFG-SRV-017\PFG-USR-UJT")
				Call MapPrinter("\\PFG-SRV-017\PFG-BRS-UMB")
												
				Log("Completed 64bit PFG-AUSTRALIS Location Printer Configuration.")
				
			End If
	End Select	
	
	'Delete old helpdesk icon.
	objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\HelpDesk.lnk")
	objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\HelpDesk.lnk")
	objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\IT Training Registration.lnk")
	objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\IT Training Registration.lnk")
		
	'Test for support folder, create if not found.
	If objFSO.FolderExists("C:\Support") = FALSE Then
		   objFSO.CreateFolder("C:\Support")
	End if

    '
    '
    'Enforce Outlook 2000 Settings
    'Outlook message arrival visual notification - ENABLED
    'objWshShell.RegWrite "HKCU\Software\Microsoft\Office\9.0\Outlook\Preferences\Notification", 1, "REG_DWORD"

    '
    '
    'Enable Outlook Administration through Exchange 2000 Server
    'Only implement if running Windows 2000
    'If strLocal_OS = "Windows 2000 Professional" And UCase(strComputerName) <> "SRV-09" And UCase(strComputerName) <> "SRV-11" Then
    '   objWshShell.Run "Regedit /S " & "\\PFGSRV-01\netlogon\EnableOutlookAdmin.reg", 7, TRUE
    'End If

  'Disable Error Handling
  On Error Goto 0

End Sub

Sub Location_PFGAU_MAIN

  'Clear Err Object
  Err.Clear

  'Enable Error Handling
  On Error Resume Next

  'Set Short Date Format
  Call Load_TNT_Regional_Settings

    'Map drives
    objNet.RemoveNetworkDrive "F:", TRUE, TRUE
    objNet.MapNetworkDrive "F:", "\\PFG-SRV-005\UserData" ,TRUE
	
	Select Case bitProcessor
		Case 32
			If Instr(1, UCase(strComputerName), "SRV") = 0 Or strComputerName = "PFGSRV-01" OR Instr(1, UCase(strUserName), "AUTO") = 0 Then
				'Map Printers
				objNet.RemovePrinterConnection "\\PFGSRV-05\PFGADMIN", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFGSRV-05\PFGASS", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFGSRV-05\PFGLOG", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFGSRV-05\PFGMAIN", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFGSRV-05\PFGMANAGE", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFGSRV-05\PFGCOLOUR", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFGSRV-05\PFGMKTG", TRUE, TRUE
				objNet.RemovePrinterConnection "\\PFGSRV-05\PFGOFFICE", TRUE, TRUE
		
				Log("Starting 32bit PFG-MTDERR Location Printer Configuration.")
				Call MapPrinter("\\PFGSRV-05\PFG-MTD-ADM")
				Call MapPrinter("\\PFGSRV-05\PFG-MTD-ASS")
				Call MapPrinter("\\PFGSRV-05\PFG-MTD-LOG")
				Call MapPrinter("\\PFGSRV-05\PFG-MTD-APC")
				Call MapPrinter("\\PFGSRV-05\PFG-MTD-MGT")
				Call MapPrinter("\\PFGSRV-05\PFG-MTD-COL")
				Call MapPrinter("\\PFGSRV-05\PFG-MTD-MKT")
				Call MapPrinter("\\PFGSRV-05\PFG-MTD-OFF")
				Log("Completed 32bit PFG-MTDERR Location Printer Configuration.")
				'objNet.AddWindowsPrinterConnection("\\PFGSRV-05\PFG-MTD-ADM")
				'objNet.AddWindowsPrinterConnection("\\PFGSRV-05\PFG-MTD-ASS")
				'objNet.AddWindowsPrinterConnection("\\PFGSRV-05\PFG-MTD-LOG")
				'objNet.AddWindowsPrinterConnection("\\PFGSRV-05\PFG-MTD-APC")
				'objNet.AddWindowsPrinterConnection("\\PFGSRV-05\PFG-MTD-MGT")
				'objNet.AddWindowsPrinterConnection("\\PFGSRV-05\PFG-MTD-COL")
				'objNet.AddWindowsPrinterConnection("\\PFGSRV-05\PFG-MTD-MKT")
				'objNet.AddWindowsPrinterConnection("\\PFGSRV-05\PFG-MTD-OFF")
			End If
		Case 64
			If Instr(1, UCase(strComputerName), "PFG-SRV-005") = 0 Then
				Log("Starting 64bit PFG-MTDERR Location Printer Configuration.")			
				Call MapPrinter("\\PFG-SRV-005\PFG-MTD-ADM")
				Call MapPrinter("\\PFG-SRV-005\PFG-MTD-APC")
				Call MapPrinter("\\PFG-SRV-005\PFG-MTD-ASS")
				Call MapPrinter("\\PFG-SRV-005\PFG-MTD-COL")
				Call MapPrinter("\\PFG-SRV-005\PFG-MTD-LOG")
				Call MapPrinter("\\PFG-SRV-005\PFG-MTD-OFF")
				Call MapPrinter("\\PFG-SRV-005\PFG-MTD-MGT")
				Call MapPrinter("\\PFG-SRV-005\PFG-MTD-MKT")
				Log("Starting 64bit PFG-MTDERR Location Printer Configuration.")							
				'objNet.AddWindowsPrinterConnection("\\PFG-SRV-005\PFG-MTD-ADM")			
				'objNet.AddWindowsPrinterConnection("\\PFG-SRV-005\PFG-MTD-APC")
				'objNet.AddWindowsPrinterConnection("\\PFG-SRV-005\PFG-MTD-ASS")
				'objNet.AddWindowsPrinterConnection("\\PFG-SRV-005\PFG-MTD-COL")
				'objNet.AddWindowsPrinterConnection("\\PFG-SRV-005\PFG-MTD-LOG")
				'objNet.AddWindowsPrinterConnection("\\PFG-SRV-005\PFG-MTD-OFF")
				'objNet.AddWindowsPrinterConnection("\\PFG-SRV-005\PFG-MTD-MGT")
				'objNet.AddWindowsPrinterConnection("\\PFG-SRV-005\PFG-MTD-MKT")
			End If
	End Select	
	
	'Delete old helpdesk icon.
	objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\HelpDesk.lnk")
	objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\HelpDesk.lnk")
	objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\IT Training Registration.lnk")
	objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\IT Training Registration.lnk")
		
	'Test for support folder, create if not found.
	If objFSO.FolderExists("C:\Support") = FALSE Then
		   objFSO.CreateFolder("C:\Support")
	End if

    '
    '
    'Enforce Outlook 2000 Settings
    'Outlook message arrival visual notification - ENABLED
    objWshShell.RegWrite "HKCU\Software\Microsoft\Office\9.0\Outlook\Preferences\Notification", 1, "REG_DWORD"

    '
    '
    'Enable Outlook Administration through Exchange 2000 Server
    'Only implement if running Windows 2000
    'If strLocal_OS = "Windows 2000 Professional" And UCase(strComputerName) <> "SRV-09" And UCase(strComputerName) <> "SRV-11" Then
    '   objWshShell.Run "Regedit /S " & "\\PFGSRV-01\netlogon\EnableOutlookAdmin.reg", 7, TRUE
    'End If

  'Disable Error Handling
  On Error Goto 0

End Sub

Sub Location_PFGAU_SERVICE

  'Clear Err Object
  Err.Clear

  'Enable Error Handling
  On Error Resume Next

   If Instr(1, UCase(strComputerName), "PFG-SRV-017") = 0 Then

    'Map drives
    objNet.RemoveNetworkDrive "G:", TRUE, TRUE
    objNet.MapNetworkDrive "G:", "\\PFG-SRV-017\PFG-FUL-DTA", TRUE

    'Logging
    Log("G: drive mapped.")

	'Map Printers
	If bitProcessor = 32 then
			
		'Map Printers
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-FUL-PPC", TRUE, TRUE
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-FUL-WS1", TRUE, TRUE
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-FUL-WS2", TRUE, TRUE
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-FUL-WAR", TRUE, TRUE
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-FUL-PPC", TRUE, TRUE
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-FUL-DSP", TRUE, TRUE
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-FUL-LOG", TRUE, TRUE
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-FUL-COL", TRUE, TRUE
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-FUL-SHP", TRUE, TRUE
		objNet.RemovePrinterConnection "\\pfgsrv-06\PFG-MIL-ADM", TRUE, TRUE
		
		'Add New Printers
		Log("Starting 32bit PFG Fulton Location Printer Configuration.")
		Call MapPrinter("\\PFG-SRV-013\PFG-FUL-COL")
		Call MapPrinter("\\PFG-SRV-013\PFG-FUL-DSP")
		Call MapPrinter("\\PFG-SRV-013\PFG-FUL-LOG")
		Call MapPrinter("\\PFG-SRV-013\PFG-FUL-WAR")
		Call MapPrinter("\\PFG-SRV-013\PFG-FUL-WS1")
		Call MapPrinter("\\PFG-SRV-013\PFG-FUL-WS2")
		Call MapPrinter("\\PFG-SRV-013\PFG-FUL-PPC")
		Call MapPrinter("\\PFG-SRV-013\PFG-FUL-ADM")
		Log("Completed 32bit PFG Fulton Location Printer Configuration.")
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-013\PFG-FUL-COL")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-013\PFG-FUL-DSP")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-013\PFG-FUL-LOG")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-013\PFG-FUL-WAR")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-013\PFG-FUL-WS1")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-013\PFG-FUL-WS2")
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-013\PFG-FUL-PPC")
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-013\PFG-MIL-ADM")
				
	ElseIf bitProcessor = 64 then
	
		'Connect New Printers
		Log("Starting 64bit PFG Fulton Location Printer Configuration.")		
		Call MapPrinter("\\PFG-SRV-017\PFG-FUL-COL")
		Call MapPrinter("\\PFG-SRV-017\PFG-FUL-DSP")
		Call MapPrinter("\\PFG-SRV-017\PFG-FUL-LOG")
		Call MapPrinter("\\PFG-SRV-017\PFG-FUL-WAR")
		Call MapPrinter("\\PFG-SRV-017\PFG-FUL-WS1")
		Call MapPrinter("\\PFG-SRV-017\PFG-FUL-WS2")
		Call MapPrinter("\\PFG-SRV-017\PFG-FUL-PPC")
		Call MapPrinter("\\PFG-SRV-017\PFG-FUL-ADM")
		Log("Completed 64bit PFG Fulton Location Printer Configuration.")
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-011\PFG-FUL-COL")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-011\PFG-FUL-DSP")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-011\PFG-FUL-LOG")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-011\PFG-FUL-WAR")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-011\PFG-FUL-WS1")	
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-011\PFG-FUL-WS2")
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-011\PFG-FUL-PPC")
		'objNet.AddWindowsPrinterConnection("\\PFG-SRV-011\PFG-MIL-ADM")		
		
	End If	 

    'Logging
    Log("Printers Connected.")

   End If

  'Standardise Regional Settings
  Call Load_TNT_Regional_Settings

    'Check OS
    Select Case strLocal_OS
       Case "Windows 95"
	  'Set system time
	  objWshShell.Run "Command /C Net Time \\192.168.201.202 /set /YES", 7, FALSE
          'Logging
          Log("Time set.")
       Case "Windows 98"
	  'Set system time
	  objWshShell.Run "Command /C Net Time \\192.168.201.202 /set /YES", 7, FALSE
          'Logging
          Log("Time set.")	  
       Case Else
	  'Set system time
	  'objWshShell.Run "Net Time \\192.168.201.202 /set /YES", 7, FALSE
     End Select
	  'objWshShell.Run "Route Add 161.71.70.204 mask 255.255.255.255 192.168.201.254", 7, FALSE
	  'objWshShell.Run "Route Add 161.71.70.200 mask 255.255.255.255 192.168.201.254", 7, FALSE
     If Instr(1, UCase(strComputerName), "SRV") = 0 Then
	  'objWshShell.Run "Route Add 192.168.203.201 mask 255.255.255.255 192.168.201.202", 7, FALSE
	  'objWshShell.Run "Route Add 192.168.202.250 mask 255.255.255.255 192.168.201.202", 7, FALSE
      End If

    'Check OS and run for Win2k systems only.
    If Instr(1, strLocal_OS, "Windows 2000") <> 0 Then

		'Log
		Log("Helpdesk Installed.")

		'Delete old helpdesk icon.
		objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\HelpDesk.lnk")
		objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\HelpDesk.lnk")
		objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\IT Training Registration.lnk")
		objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\IT Training Registration.lnk")

      End If

    'Enforce Internet Proxy Settings
    ''objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride", "192.*;168.8.152.101;txtcor02.textronturf.com*;www.anz.com;202.2.59.40;deskbank1.westpac.co.nz;deskbank2.westpac.co.nz;daedong.co.kr;anz.co.nz;coddi.com;savsystem.merlo.com;161.71.70.*;cdserver2;srv-01;srv-02;srv-03;srv-04;srv-05;srv-06;srv-07;srv-08;srv-09;srv-10;srv-11;srv-06;pfgsrv-02;pfgsrv-01;pfgsrv-06;<local>"

    '
    '
    'Enforce Outlook 2000 Settings
    'Outlook message arrival visual notification - ENABLED
    objWshShell.RegWrite "HKCU\Software\Microsoft\Office\9.0\Outlook\Preferences\Notification", 1, "REG_DWORD"

    '
    '
    'Enable Outlook Administration through Exchange 2000 Server
    'Only implement if running Windows 2000
    'If strLocal_OS = "Windows 2000 Professional" Then
    '   objWshShell.Run "Regedit /S " & "\\pfgsrv-06\netlogon\EnableOutlookAdmin.reg", 7, TRUE
    'End If

    '
    '
    'Enforce Power Policy
    'If strLocal_OS = "Windows 2000 Professional" Then
    '   objWshShell.Run "Regedit /S " & "\\pfgsrv-06\netlogon\W2klaptoppwrconf.reg", 7, TRUE
    'End If

    '
    '
    'Enforce Short Date Policy
     'objWshShell.Run "Regedit /S " & "\\pfgsrv-06\netlogon\ShortTimeFormat.reg", 7, TRUE

	'
	'
	'Enforce IDSe42 Rebranding 122004

	'Check log Key
	'strIDSe42_UPDATECODE = objWshShell.RegRead("HKLM\Software\IDS Enterprise Systems Pty Ltd\UpdateCode")
	'If Err.number <> 0 Then
	'	If Trim(strIDSe42_UPDATECODE) = "" Then
	'		'Create Key
	'		objWshShell.RegWrite "HKLM\Software\IDS Enterprise Systems Pty Ltd\UpdateCode", ""
	'		strIDSe42_UPDATECODE = "0"
	'	End If
	'End If

	'If Instr(1,strComputerName, "SRV") = 0 Then
		'strUpdate = objWshShell.Run ("\\srv-01\netlogon\PFWBRAND.hta", 1, TRUE)
	'End If

    '
    '
    'Enforce IDSe42 access string
    If (objFSO.FileExists("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.ini")) Then
	'Idse42 Software is installed. Check registry entry for mod.
	objWshShell.RegWrite "HKLM\Software\PowerFarming\", ""
	'Ini Vars
	strIDSe42_SYSTEM = ""
	'Attempt to read IDse42 installed & correctly configured Key
	strIDSe42_SYSTEM = objWshShell.RegRead("HKLM\Software\PowerFarming\IDSe42_System")

	'Test for zero length string
	If Len(strIDSe42_SYSTEM) = 0 Then
	   'Entry does not exist - create it
	     'Create RunCount Value
	     objWshShell.RegWrite "HKLM\Software\PowerFarming\IDSe42_System", "SET"

	     On Error Goto 0
	     'Rename existing file
	     objFSO.CopyFile "C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 " &_
			     "GUI\Settings.ini", "C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.OLD"
	     'DeleteFile
	     objFSO.DeleteFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.ini")
	     'Open renamed file for READING
	     Set txtIDSe42SettingsRD = objFSO.OpenTextFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.OLD", 1)
	     'Open new Settings.INI file for writing
	     Set txtIDSe42SettingsWR = objFSO.OpenTextFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.ini", 8, TRUE)

	     'Loop through items in OLD file
	     Do While txtIDSe42SettingsRD.AtEndOfStream <> True
		'ReadLine
		strCurrentLine = txtIDSe42SettingsRD.ReadLine
		'Check for system entry
		If Instr(strCurrentLine, "SYSTEM=") <> 0 Then
		   'Change the value of strCurrentline to new DNS name.
		   strCurrentLine = "SYSTEM=dealer.powerfarming.co.nz#992"
		    'Write line into new file
		    txtIDSe42SettingsWR.Write strCurrentLine
		    txtIDSe42SettingsWR.WriteLine
		Else
		    'Write line into new file
		    txtIDSe42SettingsWR.Write strCurrentLine
		    txtIDSe42SettingsWR.WriteLine
		End If
	     Loop
	End If
    End If

  'Disable Error Handling
  On Error Goto 0

End Sub

Sub Location_PFNZ_MABERS
  Err.Clear
  On Error Resume Next

	'Add WorkstationDR
	If InStr(Ucase(strComputerName), "NMAB") <> 0 Or _
		InStr(UCASE(strComputerName), "UMAB") <> 0 Or _
		InStr(UCASE(strComputerName), "PMAB") <> 0 Then	
		objWshShell.Run "Wscript.exe " & "\\powerfarming.co.nz\netlogon\svn-netlogon\login\PFW-WorkstationDR.vbs", 0, FALSE
		Log("WorkstationDR Policy Enforced")
	End If
  
    objNet.AddWindowsPrinterConnection "\\SRV-06\MMMA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\MMMB"
    objNet.AddWindowsPrinterConnection "\\SRV-06\MMMC"
    objNet.AddWindowsPrinterConnection "\\SRV-06\MMMP"
	'objNet.RemovePrinterConnection "\\SRV-06\MMMS", TRUE, TRUE	
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-MMS"

    'Call SetSystemTime
    Call DisableInternetProxy
    Call InstallSafeGuardProxyTool
    Call KillAdWatch
    Call DisableServices

	 'Create Profile
	 Call SetMAPIProfile
	 Call TightVNCInstall

    Call SetupHelpDesk
    Call GroupJobs
    Call EnforceSettings
    Call CleanUp

  'Disable Error Handling
  On Error Goto 0

End Sub

Sub GetLocation

  'Clear Err Object
  Err.Clear

  'Enable Error Handling
  On Error Resume Next

  'GetIP
  strIP = GetIP
    
  'Set Location
  If Instr(strIP, "161.71.70.") <> 0 OR Instr(strIP, "192.168.99.") <> 0 OR Instr(strIP, "192.168.50.") <> 0 OR Instr(strIP, "192.168.48.") <> 0 OR Instr(strIP, "192.168.30.") <> 0 Then
     'Set Location
     strLocation = "PFNZ"
  End If
  'Set Location
  If Instr(strIP, "192.168.208.") <> 0 OR Instr(strIP, "192.168.209.") <> 0 Then
     'Set Location
     strLocation = "PFG-AUSTRALIS"
  End If    
  'Set Location
  If Instr(strIP, "192.168.101.") <> 0 Then
     'Set Location
     strLocation = "RETAIL-TEAWAMUTU"
  End If      
  'Set Location
  If Instr(strIP, "192.168.50.197") <> 0 OR Instr(strIP, "192.168.50.196") <> 0 Then
     'Set Location
     strLocation = "PFH-RETAIL"
  End If  
  'Set Location
  If Instr(strIP, "192.168.203.") <> 0 Then
     'Set Location
     strLocation = "PFGAU_MAIN"
  End If
  'Set Location
  If Instr(strIP, "192.168.201.") <> 0 Then
     'Set Location
     strLocation = "PFGAU_SERVICE"
  End If
  'Set Location
  If Instr(strIP, "192.168.3.") <> 0 OR Instr(strIP, "192.168.109.") <> 0 Then
     'Set Location
     strLocation = "PFNZ_MABERS"
  End If
  'Set Location
  If Instr(strIP, "192.168.206.") <> 0 Then
     'Set Location
     strLocation = "PFGAU_BRISBANE"
  End If
  'Set Location
  If Instr(strIP, "192.168.0.") <> 0 Then
     'Set Location
     strLocation = "HOWARD_SYDNEY"
  End If
  'Set Location
  If Instr(strIP, "192.168.100.") <> 0 Then
     'Set Location
     strLocation = "HOWARD_SYDNEY"
  End If
  
  'Logging
  Log("Local IP Address: " & strIP)
  Log("Location: " & strLocation)

  'Disable Error Handling
  On Error Goto 0

End Sub

Sub RemoveNonLocalNetworkPrinters

  'Enable Error Handling
  On Error Resume Next

  'objWshShell.Run "Wscript.exe " & "\\powerfarming.co.nz\netlogon\RemoveNonLocalNetworkPrinters.vbs", 7, FALSE

  'Disable Error Handling
  On Error Goto 0

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
	   If Instr(strIPAddress, "192.168.201.") <> 0 _
		OR Instr(strIPAddress, "192.168.203.") <> 0 _
		OR Instr(strIPAddress, "161.71.70.") <> 0 _
		OR Instr(strIPAddress, "192.168.206.") <> 0 _
		OR Instr(strIPAddress, "192.168.3.") <> 0 _
		OR Instr(strIPAddress, "192.168.0.") <> 0 _
		OR Instr(strIPAddress, "192.168.50.") <> 0 _
		OR Instr(strIPAddress, "192.168.100.") <> 0 _
		OR Instr(strIPAddress, "192.168.48.") <> 0 _
		OR Instr(strIPAddress, "192.168.30.") <> 0 _
		OR Instr(strIPAddress, "192.168.109.") <> 0 _
		OR Instr(strIPAddress, "192.168.208.") Then
		   GetIP = strIPAddress
	   End If
	End If
     Next
  Next

  'Disable Error Handling
  On Error Goto 0

End Function

Sub DisableServices

    Err.Clear
    On Error Resume Next
	
    If strLocal_OS <> "Windows 98" AND _
		strLocal_OS <> "Windows 95" AND _
			InStr(strComputerName, "SRV") < 1 Then
			
			'Disable Webclient
			If objWshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\WebClient\Start") <> 4 Then
				strUpdate = objWshShell.Run ("net stop webclient", 0, TRUE)
			End If			
			objWshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\WebClient\Start", 4, "REG_DWORD"
	 
			'Disable OfficeScan Firewall
			If objWshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\OfcPfwSvc\Start") <> 4 Then
				strUpdate = objWshShell.Run ("net stop " & Chr(34) & "OfficescanNT Personal Firewall" & Chr(34), 0, TRUE)
			End If			
			objWshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\OfcPfwSvc\Start", 4, "REG_DWORD"
	 
			'Disable OSO Update Service
			If objWshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\OSO Update Service\Start") <> 4 Then
				strUpdate = objWshShell.Run ("net stop " & Chr(34) & "OSO Update Service" & Chr(34), 0, TRUE)
			End If
			objWshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\OSO Update Service\Start", 4, "REG_DWORD"
			
    End If
	
    On Error Goto 0
	
End Sub

Sub MSHotFix

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'This routine installs the MS hotfixes quietly.

    'This procedure has been tested on the following OS's
    '	  1. Windows 2000 Professional

    If strLocal_OS = "Windows 2000 Professional" Then

      'Check MS04-011
      strRegValCheck = objWshShell.RegRead(_
		"HKLM\SOFTWARE\Microsoft\Updates\Windows 2000\SP5\KB835732\Type")
      If strRegValCheck <> "Update" Then
	  strUpdate = objWshShell.Run ("\\srv-03\vol3\PublicSupport\MS04-011\Win2k\" & "Windows2000-KB835732-x86-ENU.EXE /quiet /norestart", 0, TRUE)
      End If

    End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GetServicePack

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'This routine gets the Computer NetBIOS Name of the local system.

    'This procedure has been tested on the following OS's
    '	  1. Windows 2000 Professional
    '	  2. Windows 2000 Server

    'Retrieve Computer NetBIOS Name
    Select Case strLocal_OS
      Case "Windows 2000 Professional"
	   'Get Service Pack Level from registry
	   strServicePack = objWshShell.RegRead(_
	      "HKLM\Software\Microsoft\Windows NT\CurrentVersion\CSDVersion")

      Case "Windows 2000 Server"
	   'Get Service Pack Level from registry
	   strServicePack = objWshShell.RegRead(_
	      "HKLM\Software\Microsoft\Windows NT\CurrentVersion\CSDVersion")
      Case Else
	strServicePack = "Unknown"

    End Select

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GetRegSvr32RunPath

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'This routine sets the strRegSvr32Path variable

    'This procedure has been tested on the following OS's
    '	  1. Windows 2000 Professional
    '	  2. Windows 2000 Server
    '	  3. Windows 98 & SE
    '	  4. Windows 95
    '	  5. Windows NT Workstation 4.0
    '	  6. Windows NT Server 4.0

    'Set strLibInstPath variable
    Select Case strLocal_OS

      Case "Windows 95"

	   'Check for existence of RegSvr32 in SYSTEM (or SYSTEM32) - set var accordingly
	   If objFSO.FileExists(strSystemRoot & "\System\Regsvr32.exe") = TRUE Then
	     'Set strLibInstPath to use SYSTEM
	     strRegSvr32Path = strSystemRoot & "\System"
	   Else
	      'Set strLibInstPath to use SYSTEM32
	      strRegSvr32Path = strSystemRoot & "\System32"
	   End If

      Case "Windows 98"

	   'Check for existence of RegSvr32 in SYSTEM (or SYSTEM32) - set var accordingly
	   If objFSO.FileExists(strSystemRoot & "\System\Regsvr32.exe") = TRUE Then
	     'Set strLibInstPath to use SYSTEM
	     strRegSvr32Path = strSystemRoot & "\System"
	   Else
	      'Set strLibInstPath to use SYSTEM32
	      strRegSvr32Path = strSystemRoot & "\System32"
	   End If

      Case "Windows NT Workstation 4.0"
	   strRegSvr32Path = strSystemRoot & "\System32"

      Case "Windows NT Server 4.0"
	   strRegSvr32Path = strSystemRoot & "\System32"

      Case "Windows 2000 Professional"
	   strRegSvr32Path = strSystemRoot & "\System32"

      Case "Windows 2000 Server"
	   strRegSvr32Path = strSystemRoot & "\System32"

    End Select

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub AutoUpdate

    'Clear Err Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Do not run for Servers
    If InStr(UCase(strComputerName), "SRV") = 0 Then

      '
      '
      'Check if IDSe42 Update is needed.
      If objFSO.GetFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\IDSE42.exe").Size <> 745472 Then
		If objFSO.GetFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\IDSRFile.ocx").Size <> 36864 Then
		strLocalIDSRFile = "c:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\IDSRFILE.OCX"
		strNetworkIDSRFile = "\\SRV-03\NETLOGON\IDSRFILE.OCX"
		'Copy and register OCX
		objFSO.CopyFile strNetworkIDSRFile, strLocalIDSRFile
			'Register DLL Server
			strDLLRegisterResult = 0
			strDLLRegisterResult = objWshShell.Run (strRegSvr32Path & "\Regsvr32 " &_
						strLocalIDSRFile & " /s", 0, TRUE)
		End If
		'If objFSO.GetFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Tnblib.ocx").Size <> 1885184 Then
		strLocalTnblibFile = "c:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\TNBLIB.OCX"
		strNetworkTnblibFile = "\\SRV-03\NETLOGON\TNBLIB.OCX"
		'Copy and register OCX
		objFSO.CopyFile strNetworkTnblibFile, strLocalTnblibFile
			'Register DLL Server
			strDLLRegisterResult = 0
			strDLLRegisterResult = objWshShell.Run (strRegSvr32Path & "\Regsvr32 " &_
						strLocalTnblibFile & " /s", 0, TRUE)
		'End If
		If objFSO.GetFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\IDSE42.exe").Size <> 745472 Then
		'Copy file only. No need to register EXE
		objFSO.CopyFile "\\SRV-03\NETLOGON\IDSE42.EXE", "c:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\IDSE42.EXE"
		End If

		'Copy file only. No need to register EXE
		objFSO.CopyFile "\\SRV-03\NETLOGON\IDSGUI.KBD", "c:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\IDSGUI.KBD"
		objFSO.CopyFile "\\SRV-03\NETLOGON\IDSGUI.TNS", "c:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\IDSGUI.TNS"
		objFSO.CopyFile "\\SRV-03\NETLOGON\LIBEAY32.DLL", "c:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\LIBEAY32.DLL"
		objFSO.CopyFile "\\SRV-03\NETLOGON\SSLEAY32.DLL", "c:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\SSLEAY32.DLL"
		objFSO.CopyFile "\\SRV-03\NETLOGON\IDSE42.EXE", "c:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\IDSE42.EXE"

		'Notify Administrator
		If Err.Number = 0 Then
			strSMTPServer = "Srv-02"
			strSMTPMailFrom = "administrator@powerfarming.co.nz"
			strSMTPSendTo = "mbarrett@powerfarming.co.nz"
			strSMTPMessageSubject = "Success Event - IDSe42 Update 2 - " & UCASE(strUserName) & " on " & UCASE(strComputerName)
			Call Mailer
		Else
			strSMTPServer = "Srv-02"
			strSMTPMailFrom = "administrator@powerfarming.co.nz"
			strSMTPSendTo = "mbarrett@powerfarming.co.nz"
			strSMTPMessageSubject = "Failure Event - IDSe42 Update 2 - " & UCASE(strUserName) & " on " & UCASE(strComputerName)
			strSMTPMessageText = Err.Number & " " & Err.Description
			Call Mailer
		End If
      End If
    End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub SetSystemTime

    'Clear Errors
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Check OS
    Select Case strLocal_OS
       Case "Windows 95"
	  'Set system time
	  objWshShell.Run "Command /C Net Time " & strLogonServer & " /set /YES", 7, FALSE
       Case "Windows 98"
	  'Set system time
	  objWshShell.Run "Command /C Net Time " & strLogonServer & " /set /YES", 7, FALSE
       Case Else
	  'Set system time
	  objWshShell.Run "Net Time " & strLogonServer & " /set /YES", 7, FALSE
    End Select

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub KillAdWatch

	'Clear Error Object
	Err.Clear

	'Disable Error Handling
	On Error Resume Next

	'Load object
	Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")

	'Init. vars
	tfCont = False

	'Init. Counter
	intCounter = 0

	'Loop until tfCont becomes True
	Do While tfCont <> True

	    'Update Counter
	    intCounter = intCounter + 1

	    'Loop through processes on the workstation and Kill and instances for ImpAdmin.exe
	    Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)
	    For Each objItem in colItems
			'Stop Impromptu if it is found to be running
			If objItem.Description = "Ad-Watch.exe" Then
				'Kill Process
				objItem.Terminate
			End If
		Next
		'Kill colItems Object
		Set colItems = Nothing

		'Init.
		tfCont = True
	    'Loop through processes on the workstation and check for instances of ImpAdmin
	    Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)
	    For Each objItem in colItems
			'Check for running instances for ImpAdmin.exe
			If objItem.Description = "Ad-Watch.exe" Then
				'Keep looping as ImpAdmin was still there
				tfCont = False
			End If
		Next

		'Kill colItems Object
		Set colItems = Nothing

	     If intCounter > 25 Then
		Exit Do
	     End If

	Loop

	'Delete Startup Reg Entries
	objWshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Ad-Watch"

	'Disable Error Handling
	On Error Goto 0

End Sub

Sub PFG_ADMIN_GroupJobs

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Log
    Log("Starting PFG_ADMIN_GroupJobs GroupJobs...")

    '
    '
    'Get user object
    Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")
    'Loop through all groups user is a member of
    For Each Group in User.Groups
		'Logging
		Log("	 User member of: " & Group.Name)
		Select Case bitProcessor
			Case 32
				If Instr(1, UCase(strComputerName), "SRV") = 0 Or strComputerName = "PFGSRV-01" OR Instr(1, UCase(strUserName), "AUTO") = 0 Then
					'Map Printers
					Select Case Group.Name
						Case "PFG_Admin_Reception_Location"
							 objNET.SetDefaultPrinter "\\PFGSRV-05\PFGMAIN"
						Case "PFG_Assembly_Location"
							 objNET.SetDefaultPrinter "\\PFGSRV-05\PFGASS"
						Case "PFG_Logistics_Location"
							 objNET.SetDefaultPrinter "\\PFGSRV-05\PFGLOG"
						Case "PFG_Management_Location"
							 objNET.SetDefaultPrinter "\\PFGSRV-05\PFGMANAGE"
						Case "PFG_Admin_Admin_Location"
							 objNET.SetDefaultPrinter "\\PFGSRV-05\PFGADMIN"
						Case "PFG_Marketing_Location"
							 objNET.SetDefaultPrinter "\\PFGSRV-05\PFGMKTG"
					End Select
				End If
			Case 64
				If Instr(1, UCase(strComputerName), "PFG-SRV-005") = 0 Then
					Select Case Group.Name
						Case "PFG_Admin_Reception_Location"
							 objNET.SetDefaultPrinter "\\PFG-SRV-005\PFG-MTD-APC"
						Case "PFG_Assembly_Location"
							 objNET.SetDefaultPrinter "\\PFG-SRV-005\PFG-MTD-ASS"
						Case "PFG_Logistics_Location"
							 objNET.SetDefaultPrinter "\\PFG-SRV-005\PFG-MTD-LOG"
						Case "PFG_Admin_Admin_Location"
							 objNET.SetDefaultPrinter "\\PFG-SRV-005\PFG-MTD-OFF"
						Case "PFG_Marketing_Location"
							 objNET.SetDefaultPrinter "\\PFG-SRV-005\PFG-MTD-COL"
					End Select			
				End If
		End Select			
    Next

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub PFG_PARTS_GroupJobs

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Log
    Log("Starting PFG_PARTS_GroupJobs GroupJobs...")

    '
    '
    'Get user object
    Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")
    'Loop through all groups user is a member of
    For Each Group in User.Groups
	'Logging
	Log("	 User member of: " & Group.Name)

	'Do work based on group membership
	Select Case Group.Name
		Case "PFG_Parts_Reception_Location"
		     objNET.SetDefaultPrinter "\\pfgsrv-06\PFGPPC"
		Case "PFG_Service_Location"
		     objNET.SetDefaultPrinter "\\pfgsrv-06\PFGPARTS"
		Case "PFG_Parts_Location"
		     objNET.SetDefaultPrinter "\\pfgsrv-06\PFGWHS1"
	End Select
	    objNet.RemovePrinterConnection "\\pfgsrv-06\PFGMKTG"
    Next

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GroupJobs

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Log
    Log("Starting GroupJobs...")

    'Set tfContinue default val
    tfTerminalUser = FALSE

    'Windows 9x will not get group memberships here unless the ADSI agent has been installed.
    If strLocal_OS <> "Windows 98" AND strLocal_OS <> "Windows 95" Then

      'Check OS and run for Citrix\Terminal servers only
      Select Case strComputerName
	  Case "SRV-09"
	       tfTerminalUser = TRUE
		   temp = CreateHomeDir(strUsername, "\\pfnz-srv-028\PFWData\Home\")
	       'temp = CreateHomeDir(strUsername, "\\srv-01\vol1\home\")
	  Case "PFNZ-SRV-002"
	       tfTerminalUser = TRUE
	       temp = CreateHomeDir(strUsername, "\\pfnz-srv-028\PFWData\Home\")
	  Case "PFNZ-SRV-032"
	       tfTerminalUser = TRUE
	       temp = CreateHomeDir(strUsername, "\\pfnz-srv-028\PFWData\Home\")		   
	  Case "SRV-11"
	       tfTerminalUser = TRUE
	       temp = CreateHomeDir(strUsername, "\\pfnz-srv-028\PFWData\Home\")
	  Case "PFNZ-IT-004"
			tfTerminalUser = TRUE
	  Case "PFNZ-SRV-035"
			tfTerminalUser = TRUE	  
	  Case "PFNZ-SRV-038"
			tfTerminalUser = TRUE	  			
			'temp = CreateHomeDir(strUsername, "\\pfnz-srv-028\mydocs$\") --> now controlled in GPO
			'temp = CreateHomeDir(strUsername, "\\pfnz-srv-028\usrprof$\") --> now controlled in GPO
      End Select

      'Logging
      'If tfTerminalUser = TRUE Then
		'Log("System is a terminal server.")
      'Else
		'Log("System is Workstation / Laptop.")
      'End If

      'Check continuance
      If tfTerminalUser = FALSE Then

	      'Call RemoveNonLocalNetworkPrinters
		  Call CreateAXShortcuts

			'
			'
			'This user is NOT logged onto one of the previously listed terminal server
			'Get user object
			Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")
			'Loop through all groups user is a member of
			For Each Group in User.Groups
			    'Logging
			    Log("    User member of: " & Group.Name)

			    'Do work based on group membership
			    Select Case Group.Name
				    Case "CD Server User"
						Call CDSERVER_DRIVE_MAP
				    Case "mapOtoRETDATA"
						Call GROUPJOB_mapOtoRETDATA
				    Case "PFBI"
						Call GROUPJOB_PFBI
				    Case "AS400 Share Connect"
						Call GROUPJOB_AS400_Share_Connect
					Case "pfnz_marketing"
						Call GROUPJOB_PFNZ_MARKETING
					Case "DemandSolutionsPFGPFW"
						Call DemandSolutionsPFGPFW
					Case "DemandSolutionsHAU"
						Call DemandSolutionsPFGPFW
					Case "PFG_AX_Service_RemoteAppServer_Drive"
						Call PFG_AX_Service_RemoteAppServer_Drive
			    End Select
			Next

			If FamisInstalled = True Then
			
				'Map require drive
				If objNet.FolderExists("U:") = True Then
				   Err.Clear
				Log("	   U: drive found.. attempting to disconnect.")
				objNet.RemoveNetworkDrive "U:", TRUE, TRUE
				Do While objNet.FolderExists("U:")
					If objNet.FolderExists("U:") = False Then
					   Exit Do
					End If
					Wscript.Sleep 1000
					intCounter = intCounter + 1
					If intCounter = 10 Then
					   Exit Do
					End If
				Loop
				End If

				If objNet.FolderExists("U:") = False Then
				  objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA", FALSE
				  If objNet.FolderExists("U:") = True Then
				Log("	   U: drive mapped successfully.")
				  Else
				   Log("      Unable to connect U: drive.(" & Err.Number & ", " & Err.Description & ")")
				  End If
				End If			
			
				For Each Group in User.Groups					
					'Do work based on group membership
					Select Case Group.Name
						Case "stdgrpAshburton"
						 'Call GROUPJOB_stdgrp_ASHBURTON
						Case "stdgrpAuckland"
						 'Call GROUPJOB_stdgrp_AUCKLAND
						Case "stdgrp_WestCoast"
						 'Call GROUPJOB_stdgrp_WestCoast
						Case "stdgrp_MABERMOTORS"
						 'Call GROUPJOB_stdgrp_MABERMOTORS
						Case "stdgrp_AgriLife"
						 'Call GROUPJOB_stdgrp_AgriLife
						Case "stdgrp_Training"
						 'Call GROUPJOB_stdgrp_Training
						Case "stdgrp_PowerTrac"
						 'Call GROUPJOB_stdgrp_POWERTRAC
						Case "stdgrp_PFROTORUA"
						 'Call GROUPJOB_stdgrp_PFROTORUA
						Case "stdgrp_PFMANAWATU"
						 'Call GROUPJOB_stdgrp_PFMANAWATU
						Case "stdgrp_PFAGSOUTHLAND"
						 'Call GROUPJOB_stdgrp_AGSOUTHLAND
						Case "stdgrp_MABERTRACTORS"
						 'Call GROUPJOB_stdgrp_MABERTRACTORS
						Case "stdgrp_AgEarth"
						 'Call GROUPJOB_stdgrp_AGEARTH
						Case "stdgrp_BROWNWOODS"
						 'Call GROUPJOB_stdgrp_BROWNWOODS
						Case "stdgrp_CanterburyTractors"
						 'Call GROUPJOB_stdgrp_CANTERBURY
						Case "stdgrp_PREMIER"
						 'Call GROUPJOB_stdgrp_PREMIER
						Case "stdgrpOTAGO"
						 'Call GROUPJOB_stdgrp_OTAGO
						Case "stdgrp_TTCHire"
						 'Call GROUPJOB_stdgrp_TTCHire
						Case "stdgrp_GISBORNE"
						 'Call GROUPJOB_stdgrp_GISBORNE
						Case "stdgrp_HAMILTON"
						 'Call GROUPJOB_stdgrp_HAMILTON
						Case "stdgrp_PFGPowerfurf"
						 'Call GROUPJOB_stdgrp_PFGPowerTurf
						Case "stdgrp_HowardEngineering"
						  Call GROUPJOB_stdgrp_HowardEngineering
					End Select			    
				Next							
			End If
    Else
	
		'Retail project to move to 64 bit servers.
		If strComputerName = "SRV-09" or _
			strComputerName = "PFNZ-SRV-002" or _
				strComputerName = "PFNZ-SRV-032" or _
					strComputerName = "SRV-11" Then
					
			Log("	   User logged into 32bit Retail Environment.")					
					
			'
			'
			'This user is logged onto one of the Citrix\Terminal servers

			'Make Client Access Registry Mods
			objWshShell.RegDelete "HKCU\Software\Microsoft\Office\9.0\Excel\Options\Open"
			objWshShell.RegDelete "HKCU\Software\Microsoft\Office\9.0\Excel\Options\Open1"
			objWshShell.RegDelete "HKCU\Software\Microsoft\Office\9.0\Excel\Options\Open2"

			'Get user object
			Err.Clear
			Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")

			'Some drive cleanups
			'objNet.RemoveNetworkDrive "X:", True, True			
			
				'Map FAMIS required drive
				If objNet.FolderExists("U:") = True Then
				   Err.Clear
					Log("	   U: drive found.. attempting to disconnect.")
					objNet.RemoveNetworkDrive "U:", TRUE, TRUE
				Do While objNet.FolderExists("U:")
					If objNet.FolderExists("U:") = False Then
					   Exit Do
					End If
					Wscript.Sleep 1000
					intCounter = intCounter + 1
					If intCounter = 10 Then
					   Exit Do
					End If
				Loop
				End If

				If objNet.FolderExists("U:") = False Then
				  objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA", FALSE
				  If objNet.FolderExists("U:") = True Then
				Log("	   U: drive mapped successfully.")
				  Else
				   Log("      Unable to connect U: drive.(" & Err.Number & ", " & Err.Description & ")")
				  End If
				End If				
			
			'Loop through all groups user is a member of
			For Each Group in User.Groups
			Log("	 User member of: " & Group.Name)

			'Cumulative Groupings - Do work based on group membership
			Select Case Group.Name
				Case "stdgrp_AgriLife"
				'Call GROUPJOB_stdgrp_AgriLife
				Case "stdgrp_Training"
				'Call GROUPJOB_stdgrp_Training
				Case "stdgrp_PowerTrac"
				'Call GROUPJOB_stdgrp_POWERTRAC
				Case "stdgrp_PFROTORUA"
				'Call GROUPJOB_stdgrp_PFROTORUA
				Case "stdgrp_PFMANAWATU"
				'Call GROUPJOB_stdgrp_PFMANAWATU
				Case "stdgrp_PFAGSOUTHLAND"
				'Call GROUPJOB_stdgrp_AGSOUTHLAND
				Case "mapOtoRETDATA"
				'Call GROUPJOB_mapOtoRETDATA
				Case "stdgrp_WestCoast"
				'Call GROUPJOB_stdgrp_WestCoast
				Case "stdgrp_MABERMOTORS"
				'Call GROUPJOB_stdgrp_MABERMOTORS
				Case "stdgrpAshburton"
				'Call GROUPJOB_stdgrp_ASHBURTON
				Case "stdgrpAuckland"
				'Call GROUPJOB_stdgrp_AUCKLAND
				Case "stdgrp_MABERTRACTORS"
				'Call GROUPJOB_stdgrp_MABERTRACTORS
				Case "stdgrp_AgEarth"
				'Call GROUPJOB_stdgrp_AGEARTH
				Case "PFC1"
				'Call GROUPJOB_PFC1
				Case "stdgrp_BROWNWOODS"
				'Call GROUPJOB_stdgrp_BROWNWOODS
				Case "stdgrp_CanterburyTractors"
				'Call GROUPJOB_stdgrp_CANTERBURY
				Case "stdgrp_PREMIER"
				'Call GROUPJOB_stdgrp_PREMIER
				Case "AS400 Share Connect"
				'Call GROUPJOB_AS400_Share_Connect
				Case "stdgrpOTAGO"
				'Call GROUPJOB_stdgrp_OTAGO
				Case "stdgrp_TTCHire"
				'Call GROUPJOB_stdgrp_TTCHire
				Case "stdgrp_GISBORNE"
				'Call GROUPJOB_stdgrp_GISBORNE
				Case "stdgrp_HAMILTON"
				'Call GROUPJOB_stdgrp_HAMILTON
				Case "PFBI"
				'Call GROUPJOB_PFBI
				Case "stdgrp_PFGPowerturf"
				'Call GROUPJOB_stdgrp_PFGPowerTurf
				Case "localts_MORRINSVILLE"
				'Call localts_MORRINSVILLE
				Case "stdgrp_HowardEngineering"
				 Call GROUPJOB_stdgrp_HowardEngineering
			End Select

			'Printer Groupings
			Select Case Group.Name
			       Case "P31"
				   objNet.AddWindowsPrinterConnection("\\SRV-01\P31")
			End Select
			Next
		ElseIf strComputerName = "PFNZ-SRV-035" Then
			Log("	   User logged into 64bit Retail Environment.")					
			Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")			
			For Each Group in User.Groups
				Log("	 User member of: " & Group.Name)			
				Select Case Group.Name			
					Case "stdgrp_MABERMOTORS"
						Call GROUPJOB_stdgrp_MABERMOTORS_2012
				End Select
			Next
		End If
      End If

    ElseIf strLocal_OS = "Windows 98" Then

	'
	'
	'Run Office Scan Update
	 'objWshShell.Run "\\srv-03\ofcscan\AUTOPCC.EXE", 7, FALSE
	 ''Map CD Server Drives
	 'Call CDSERVER_DRIVE_MAP

    End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub DisableInternetProxy

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Clear Proxy Settings
	'objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
	'objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", ""
	'objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride", ""

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub InstallSMTP

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

	'Check SafeGuard file and copy
	'If objFSO.FileExists(strSystemRoot & "\smtp.ocx") = False Then
	       objFSO.CopyFile "\\srv-03\netlogon\smtp.ocx", strSystemRoot & "\system32\"
	       'Wscript.Echo strSystemRoot & "\system32\regsvr32 /s smtp.ocx"
	      strUpdate = objWshShell.Run (strSystemRoot & "\system32\regsvr32 smtp.ocx /s", 0, TRUE)

	'Else
	      'strUpdate = objWshShell.Run (strSystemRoot & "\system32\regsvr32 smtp.ocx /s", 0, TRUE)
	'End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub InstallSafeGuardProxyTool

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

	'Wscript.Echo strSystemRoot & "\safeguard.exe"

	'Check SafeGuard file and copy
	If objFSO.FileExists(strSystemRoot & "\safeguard.exe") = False Then
	       'objFSO.CopyFile "\\powerfarming.co.nz\netlogon\safeguard.exe", strSystemRoot & "\"
	      'strUpdate = objWshShell.Run (strSystemRoot & "\safeguard.exe", 0, TRUE)
	      'objWshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\SafeGuardProxySet", strSystemRoot & "\safeguard.exe"
	      'Log("SafeGuard.exe was added to the run reg and run.")

	Else
              objFSO.DeleteFile strSystemRoot & "\safeguard.exe"
	      Log("SafeGuard.exe was found, delete was attempted.")
	End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_TTCHire

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_TTCHire")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\TTCHire$"
    End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_OTAGO

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_OTAGO")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PFOTAGO$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection("\\SRV-06\PFOS")
    objNet.RemovePrinterConnection "\\SRV-06\PFOP", TRUE, TRUE
    objNet.AddWindowsPrinterConnection("\\SRV-06\PFOC")
    objNet.AddWindowsPrinterConnection("\\SRV-06\PFOA")
	
    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Otago.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Otago.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfotago"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Otago Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save			    
	
    'Disable Error Handling
    On Error Goto 0

End Sub


Sub GROUPJOB_stdgrp_HowardEngineering

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_HowardEngineering")

    'Map drives
    objNet.RemoveNetworkDrive "J:", True, True
    objNet.RemoveNetworkDrive "K:", True, True
    objNet.RemoveNetworkDrive "O:", True, True
    objNet.RemoveNetworkDrive "U:", True, True
    objNet.RemoveNetworkDrive "X:", True, True
    objNet.RemoveNetworkDrive "I:", True, True
    objNet.RemoveNetworkDrive "M:", True, True
    'objNet.MapNetworkDrive "M:", "\\SRV-01\HENG"

	'Cutover from SRV-01 to PFNZ-SRV-028 on 21/11/2013
	On Error Resume Next
	If objNet.FolderExists("M:") = True Then
	   Err.Clear
		Log("M: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "M:", TRUE, TRUE
		Do While objNet.FolderExists("M:")
			If objNet.FolderExists("M:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If		
	objNet.MapNetworkDrive "M:", "\\PFNZ-SRV-028\HENG"
	
    objNet.MapNetworkDrive "N:", "\\SRV-14\SYSPRO$"	
	Call CDSERVER_DRIVE_MAP	

    'Map Printers
	objNet.RemovePrinterConnection "\\SRV-01\HENGA", TRUE, TRUE	
	objNet.RemovePrinterConnection "\\SRV-01\HENGB", TRUE, TRUE	
	objNet.RemovePrinterConnection "\\SRV-01\HENGC", TRUE, TRUE	
	objNet.RemovePrinterConnection "\\SRV-01\HENGD", TRUE, TRUE	
	objNet.RemovePrinterConnection "\\SRV-01\HENGE", TRUE, TRUE	
	objNet.RemovePrinterConnection "\\SRV-01\HENGF", TRUE, TRUE	
	objNet.RemovePrinterConnection "\\SRV-01\HENGG", TRUE, TRUE	
	objNet.RemovePrinterConnection "\\SRV-01\HENGH", TRUE, TRUE		
	objNet.RemovePrinterConnection "\\SRV-01\HENGI", TRUE, TRUE			
			
    Call MapPrinter("\\PFNZ-SRV-029\HENGA")
    Call MapPrinter("\\PFNZ-SRV-029\HENGB")
    Call MapPrinter("\\PFNZ-SRV-029\HENGC")
    Call MapPrinter("\\PFNZ-SRV-029\HENGD")
    Call MapPrinter("\\PFNZ-SRV-029\HENGE")
    Call MapPrinter("\\PFNZ-SRV-029\HENGG")
    Call MapPrinter("\\PFNZ-SRV-029\HENGH")
    Call MapPrinter("\\PFNZ-SRV-029\HENGI")


	'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\HE SysPro 6.0.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\HE SysPro 6.0.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\Program Files\SYSPRO60\base\IMPCSC.EXE"
	objHelpdeskDesktopLink.Arguments = "/HOST=syspro.powerfarming.co.nz"
	objHelpdeskDesktopLink.WindowStyle = 1
	'objHelpdeskDesktopLink.IconLocation = "%ProgramFiles%\SYSPRO60\Renault.ico"
	objHelpdeskDesktopLink.Description = "SysPro 6.0"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\Program Files\SYSPRO60\base"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_HowardEngineering_2014

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_HowardEngineering_2014")

    'Map drives
    objNet.RemoveNetworkDrive "J:", True, True
    objNet.RemoveNetworkDrive "K:", True, True
    objNet.RemoveNetworkDrive "O:", True, True
    objNet.RemoveNetworkDrive "U:", True, True
    objNet.RemoveNetworkDrive "X:", True, True
    objNet.RemoveNetworkDrive "I:", True, True
    objNet.RemoveNetworkDrive "M:", True, True
    'objNet.MapNetworkDrive "M:", "\\SRV-01\HENG"

	'Cutover from SRV-01 to PFNZ-SRV-028 on 21/11/2013
	On Error Resume Next
	If objNet.FolderExists("M:") = True Then
	   Err.Clear
		Log("M: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "M:", TRUE, TRUE
		Do While objNet.FolderExists("M:")
			If objNet.FolderExists("M:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop	
	End If		
	objNet.MapNetworkDrive "M:", "\\PFNZ-SRV-028\HENG"	
    objNet.MapNetworkDrive "N:", "\\SRV-14\SYSPRO$"	
	Call CDSERVER_DRIVE_MAP	

    'Map Printers			
    Call MapPrinter("\\PFNZ-SRV-028\HENGA")
    Call MapPrinter("\\PFNZ-SRV-028\HENGB")
    Call MapPrinter("\\PFNZ-SRV-028\HENGC")
    Call MapPrinter("\\PFNZ-SRV-028\HENGD")
    Call MapPrinter("\\PFNZ-SRV-028\HENGE")
    Call MapPrinter("\\PFNZ-SRV-028\HENGG")
    Call MapPrinter("\\PFNZ-SRV-028\HENGH")
    Call MapPrinter("\\PFNZ-SRV-028\HENGI")

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\SysPro.lnk")
	objFSO.CopyFile "\\powerfarming.co.nz\netlogon\svn-netlogon\RemoteAppRdpShortcuts\FullColour\SysPro.lnk", strDesktop & "\SysPro.lnk"

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_PREMIER

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_PREMIER")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PREMIER$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFCS"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFCP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFCA"
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-TARA"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Premier.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Premier.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pftaranaki"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Premier Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub


Sub GROUPJOB_stdgrp_CANTERBURY

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next
    
    Log("     GROUPJOB_stdgrp_CANTERBURY")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PFCANTERBURY$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\CANA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\CANP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\CANS"
	objNet.AddWindowsPrinterConnection "\\SRV-06\CANO"
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-CCA"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Canterbury.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Canterbury.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfcanterbury"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Canterbury Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_BROWNWOODS

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("    GROUPJOB_stdgrp_BROWNWOODS")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PFTIMARU$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\BWTA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\BWTP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\BWTS"	  
    objNet.AddWindowsPrinterConnection "\\SRV-06\BWTK"
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-PFTG"	

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Timaru.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Timaru.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pftimaru"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "BrownWoods Timaru Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_ASHBURTON

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_ASHBURTON")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PFAshburton$"
    End If

    'Map Printers
	objNet.RemovePrinterConnection("\\SRV-06\ASHA")
	objNet.RemovePrinterConnection("\\SRV-06\ASHP")
	objNet.RemovePrinterConnection("\\SRV-06\ASHS")
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-PFAA"
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-PFAP"
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-PFAS"

    'Create ShortCuts
	'strDesktop = objWshShell.SpecialFolders("Desktop")
	'objFSO.DeleteFile(strDesktop & "\F2 Ashburton.lnk")
	'Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\F2 Ashburton.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfashburton"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Power Farming Ashburton Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0

End Sub


Sub GROUPJOB_stdgrp_TEAWAMUTU_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_TEAWAMUTU_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")	
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFTEAWAMUTUBRANCHDATA"
    End If	
	
    'Map Printers	
	objNet.RemovePrinterConnection "\\PFNZ-SRV-034\RET-AWA-ADM", TRUE, TRUE	
	objNet.RemovePrinterConnection "\\PFNZ-SRV-034\RET-AWA-PRT", TRUE, TRUE	
	objNet.RemovePrinterConnection "\\PFNZ-SRV-034\RET-AWA-SVC", TRUE, TRUE	
    Call MapPrinter("\\PFNZ-SRV-028\RET-AWA-ADM")
    Call MapPrinter("\\PFNZ-SRV-028\RET-AWA-ADM2")
	Call MapPrinter("\\PFNZ-SRV-028\RET-AWA-PRT")
	Call MapPrinter("\\PFNZ-SRV-028\RET-AWA-SVC")

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Te Awamutu.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Te Awamutu.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfwaikato"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Te Awamutu Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_TIMARU_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_TIMARU_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFTIMARUBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-TIM-ADM")
    Call MapPrinter("\\PFNZ-SRV-028\RET-TIM-PRT")
	Call MapPrinter("\\PFNZ-SRV-028\RET-TIM-COP")
	Call MapPrinter("\\PFNZ-SRV-028\RET-TIM-SVC")
	
	'Detect Mult User F2 Required
	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")			
	For Each Group in User.Groups		
		If Group.Name = "sec_Timaru.Counter" Then
			F2RestrictedLogon = True
		End If
	Next	

	If F2RestrictedLogon = True Then		
	
		Log("     	Alternate multi-user F2 enforced login application set.")		
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Timaru.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Timaru.lnk")
		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\Program Files (x86)\BHQSoftware\BHQ F2 Login\BHQF2Login.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pftimaru"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
		objHelpdeskDesktopLink.Description = "Timaru Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save		
	
	Else
	
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Timaru.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Timaru.lnk")

		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pftimaru"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\Same.ico"
		objHelpdeskDesktopLink.Description = "Timaru Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save		

	End If

    'Disable Error Handling
    On Error Goto 0
	
End Sub

Sub GROUPJOB_stdgrp_GORE_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_GORE_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFGOREBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-GOR-ADM")
	'Call MapPrinter("\\PFNZ-SRV-028\RET-GOR-PT1")
	'Call MapPrinter("\\PFNZ-SRV-028\RET-GOR-PT2")
	Call MapPrinter("\\PFNZ-SRV-028\RET-GOR-SVC")

	'Detect Mult User F2 Required
	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")			
	For Each Group in User.Groups		
		If Group.Name = "sec_Gore.Counter" Then
			F2RestrictedLogon = True
		End If
	Next		
	
	If F2RestrictedLogon = True Then	
	
		Log("     	Alternate multi-user F2 enforced login application set.")		
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Gore.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Gore.lnk")
		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\Program Files (x86)\BHQSoftware\BHQ F2 Login\BHQF2Login.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfgore"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
		objHelpdeskDesktopLink.Description = "Gore Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save			
	
	Else
	
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Gore.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Gore.lnk")
		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfgore"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
		objHelpdeskDesktopLink.Description = "Gore Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save			    
		
	End If

    'Disable Error Handling
    On Error Goto 0
	
End Sub

Sub GROUPJOB_stdgrp_INVERCARGILL_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_INVERCARGILL_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFINVERCARGILLBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-INV-ADM")
	Call MapPrinter("\\PFNZ-SRV-028\RET-INV-ADM2")
	Call MapPrinter("\\PFNZ-SRV-028\RET-INV-COL")
	Call MapPrinter("\\PFNZ-SRV-028\RET-INV-PRT")
	Call MapPrinter("\\PFNZ-SRV-028\RET-INV-SVC")
	Call MapPrinter("\\PFNZ-SRV-028\RET-INV-WRK")

	'Detect Mult User F2 Required
	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")			
	For Each Group in User.Groups		
		If Group.Name = "sec_Invercargill.Counter" Then
			F2RestrictedLogon = True
		End If
	Next	
	
	If F2RestrictedLogon = True Then
	
		Log("     	Alternate multi-user F2 enforced login application set.")

		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Invercargill.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Invercargill.lnk")

		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\Program Files (x86)\BHQSoftware\BHQ F2 Login\BHQF2Login.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfsouthland"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
		objHelpdeskDesktopLink.Description = "Invercargill Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save					
	
	Else	
	
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Invercargill.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Invercargill.lnk")

		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfsouthland"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
		objHelpdeskDesktopLink.Description = "Invercargill Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save			    

	End If
		
    'Disable Error Handling
    On Error Goto 0


End Sub

Sub GROUPJOB_stdgrp_OTAGO_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_OTAGO_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFOTAGOBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-ADM")
	Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-COL")
	Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-SVC")
	'Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-URB")
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-OTA-URB", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-BAL-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-BAL-SVC", TRUE, TRUE
	'Call MapPrinter("\\PFNZ-SRV-028\RET-BAL-ADM")
	'Call MapPrinter("\\PFNZ-SRV-028\RET-BAL-SVC")
	Call MapPrinter("\\PFNZ-SRV-028\RET-OTA-SAL")
	
	sec_PartsCounter = False
	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")			
	For Each Group in User.Groups
		If Group.Name = "sec_PartsCounter" Then
			sec_PartsCounter = True
		Else 
			
		End If
	Next
	
	If sec_PartsCounter = True Then
	
	Log("     	User is a member of sec_PartsCounter. Alternate multi-user F2 application set.")
	
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Otago.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Otago.lnk")

		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\Program Files (x86)\BHQSoftware\BHQ F2 Login\BHQF2Login.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfOtago"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
		objHelpdeskDesktopLink.Description = "Otago Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save			
	Else
    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Otago.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Otago.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfOtago"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Otago Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			
    
	End if
	
    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_CANTERBURY_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_CANTERBURY_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFCANTERBURYBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-CAN-ADM")
	Call MapPrinter("\\PFNZ-SRV-028\RET-CAN-OFF")    
	Call MapPrinter("\\PFNZ-SRV-028\RET-CAN-SVC")
	Call MapPrinter("\\PFNZ-SRV-028\RET-CAN-SVC2")    
	Call MapPrinter("\\PFNZ-SRV-028\RET-CAN-PRT")    

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Canterbury.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Canterbury.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfcanterbury"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Canterbury Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_ASHBURTON_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_ASHBURTON_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFASHBURTONBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-ASH-ADM")
	Call MapPrinter("\\PFNZ-SRV-028\RET-ASH-PRT")    
	Call MapPrinter("\\PFNZ-SRV-028\RET-ASH-SVC")
	Call MapPrinter("\\PFNZ-SRV-028\RET-ASH-UBA")	
	
	sec_PartsCounter = False
	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")			
	For Each Group in User.Groups
		If Group.Name = "sec_PartsCounter" Then
			sec_PartsCounter = True
		Else 
			
		End If
	Next
	
	If sec_PartsCounter = True Then
	
	Log("     	User is a member of sec_PartsCounter. Alternate multi-user F2 application set.")
	
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Ashburton.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Ashburton.lnk")

		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\Program Files (x86)\BHQSoftware\BHQ F2 Login\BHQF2Login.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfashburton"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
		objHelpdeskDesktopLink.Description = "Ashburton Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save			
	Else
    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Ashburton.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Ashburton.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfashburton"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Ashburton Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    
	
	End If
    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_WESTCOAST_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_WESTCOAST_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFWESTCOASTBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-WST-ADM")

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\WestCoast.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\WestCoast.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfwestcoast"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "West Coast Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0
	
End Sub

Sub GROUPJOB_stdgrp_TARANAKI_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_TARANAKI_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFTARANAKIBRANCHDATA"
    End If	
	
    'Map Printers
	'Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-ADM")
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TAR-ADM", TRUE, TRUE
	Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-ADM2")
    Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-PRT")
	Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-SVC")
	Call MapPrinter("\\PFNZ-SRV-028\RET-TAR-SAL")

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Taranaki.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Taranaki.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pftaranaki"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Taranaki Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0
	
End Sub

Sub GROUPJOB_stdgrp_MANAWATU_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_MANAWATU_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFMANAWATUBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-MWT-ADM")
    Call MapPrinter("\\PFNZ-SRV-028\RET-MWT-PRT")
	Call MapPrinter("\\PFNZ-SRV-028\RET-MWT-SVC")

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Manawatu.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Manawatu.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfmanawatu"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Manawatu Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0
	
End Sub

Sub GROUPJOB_stdgrp_HAWKESBAY_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_HAWKESBAY_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFHAWKESBAYBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-HWK-ADM")
    Call MapPrinter("\\PFNZ-SRV-028\RET-HWK-PRT")
	Call MapPrinter("\\PFNZ-SRV-028\RET-HWK-COL")
	Call MapPrinter("\\PFNZ-SRV-028\RET-HWK-KON")

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\HawkesBay.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\HawkesBay.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfhawkesbay"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Hawkes Bay Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0
	
End Sub

Sub GROUPJOB_stdgrp_GISBORNE_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_GISBORNE_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFGISBORNEBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-GIS-ADM")
    Call MapPrinter("\\PFNZ-SRV-028\RET-GIS-PRT")
	Call MapPrinter("\\PFNZ-SRV-028\RET-GIS-ADM2")

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Gisborne.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Gisborne.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfgisborne"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Gisborne Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0
	
End Sub

Sub GROUPJOB_stdgrp_AGRILIFE_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_AGRILIFE_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\AGRILIFEBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-AGL-ADM")
    Call MapPrinter("\\PFNZ-SRV-028\RET-AGL-PRT")

	'Detect Mult User F2 Required
	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")			
	For Each Group in User.Groups		
		If Group.Name = "sec_Agrilife.Counter" Then
			F2RestrictedLogon = True
		End If
	Next
	
	If F2RestrictedLogon = True Then		

		Log("     	Alternate multi-user F2 enforced login application set.")		
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\AgriLife.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\AgriLife.lnk")
		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\Program Files (x86)\BHQSoftware\BHQ F2 Login\BHQF2Login.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2agrilife"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
		objHelpdeskDesktopLink.Description = "AgriLife Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
		objHelpdeskDesktopLink.Save			    		
	
	Else
	
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\AgriLife.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\AgriLife.lnk")
		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2agrilife"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
		objHelpdeskDesktopLink.Description = "AgriLife Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
		objHelpdeskDesktopLink.Save			    
	
	End If

    'Disable Error Handling
    On Error Goto 0
	
End Sub

Sub GROUPJOB_stdgrp_NORTHLAND_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_NORTHLAND_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFNORTHLANDBRANCHDATA"
    End If	
	
    'Map Printers
	objNet.RemovePrinterConnection "\\PPFW088\RET-DAR-COM", TRUE, TRUE
	Call MapPrinter("\\PFNZ-SRV-028\RET-NTH-ADM")
	Call MapPrinter("\\PFNZ-SRV-028\RET-NTH-OFF")
	Call MapPrinter("\\PFNZ-SRV-028\RET-NTH-SVC")
	Call MapPrinter("\\PFNZ-SRV-028\RET-DAR-ADM")
	Call MapPrinter("\\PFNZ-SRV-028\RET-DAR-PRT")
	Call MapPrinter("\\PFNZ-SRV-028\RET-DAR-SVC")

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\PF Northland.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\PF Northland.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfwhangarei"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "PF Northland Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_MABERMOTORS_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_MABERMOTORS_2012")
	
    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")	
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\MABERMOTORSBRANCHDATA"
    End If	
	
    'Map Printers
	Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-ADM")
	Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-ADM2")
	Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-PRT")
	Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-PRT2")
	Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-SVC")
	Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-SVC2")
	Call MapPrinter("\\PFNZ-SRV-028\RET-MMM-UGH")
	
	'Detect Multi User F2 Required
	sec_PartsCounter = False
	Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")			
	For Each Group in User.Groups
		If Group.Name = "sec_PartsCounter" Then
			sec_PartsCounter = True
		Else 
			
		End If
	Next
	
	If sec_PartsCounter = True Then
	
		Log("     	User is a member of sec_PartsCounter. Alternate multi-user F2 application set.")
	
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Maber Motors.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Maber Motors.lnk")

		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\Program Files (x86)\BHQSoftware\BHQ F2 Login\BHQF2Login.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfmabermotors"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
		objHelpdeskDesktopLink.Description = "Maber Motors Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save			
	
	Else
	
		'Create ShortCuts
		strDesktop = objWshShell.SpecialFolders("Desktop")
		objFSO.DeleteFile(strDesktop & "\Maber Motors.lnk")
		Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Maber Motors.lnk")

		'Create Helpdesk Desktop Shortcut
		objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
		objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfmabermotors"
		objHelpdeskDesktopLink.WindowStyle = 1
		objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
		objHelpdeskDesktopLink.Description = "Maber Motors Famis 2"
		objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
		objHelpdeskDesktopLink.Save		

	End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_MABERMOTORS

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_MABERMOTORS")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\MABERMOTORS$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\MMMA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\MMMB"
    objNet.AddWindowsPrinterConnection "\\SRV-06\MMMC"
    objNet.AddWindowsPrinterConnection "\\SRV-06\MMMP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\MMMS"
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-MMS"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Maber Motors.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Maber Motors.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfmabermotors"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Maber Motors Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_AUCKLAND_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_AUCKLAND_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")	
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFAUCKLANDBRANCHDATA"
    End If	
	
    'Map Printers	
	objNet.RemovePrinterConnection "\\PFNZ-SRV-034\RET-AKL-ADM", TRUE, TRUE	
    Call MapPrinter("\\PFNZ-SRV-028\RET-AKL-ADM")


    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Auckland.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Auckland.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfauckland"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Auckland Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_WAIRARAPA_2012

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_WAIRARAPA_2012")

    'Map drives
    objNet.RemoveNetworkDrive "U:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map U: drive.")
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA"
    End If
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
	  Log("       Attempting to map P: drive.")	
      objNet.MapNetworkDrive "P:", "\\PFNZ-SRV-028\PFWAIRARAPABRANCHDATA"
    End If	
	
    'Map Printers	
	objNet.RemovePrinterConnection "\\PFNZ-SRV-034\RET-WRP-ADM", TRUE, TRUE	
    Call MapPrinter("\\PFNZ-SRV-028\RET-WRP-ADM")


    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Wairarapa.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Wairarapa.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfwairarapa"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\dfRenault.ico"
	objHelpdeskDesktopLink.Description = "Wairarapa Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0

End Sub


Sub GROUPJOB_stdgrp_PFMANAWATU

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next
    
    Log("     GROUPJOB_stdgrp_PFMANAWATU")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PFMANAWATU$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\MWTA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\MWTP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\MWTS"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Manawatu.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Manawatu.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfmanawatu"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Manawatu Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save			    

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_mapOtoRETDATA

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    If objNet.FolderExists("O:") = True Then
       Err.Clear
	Log("	   O: drive found.. attempting to disconnect.")
	objNet.FolderExists "O:", TRUE, TRUE
	Do While objNet.FolderExists("O:")
	    If objNet.FolderExists("O:") = False Then
	       Exit Do
	    End If
	    Wscript.Sleep 1000
	    intCounter = intCounter + 1
	    If intCounter = 10 Then
	       Exit Do
	    End If
	Loop
    End If

    If objNet.FolderExists("O:") = False Then
      objNet.MapNetworkDrive "O:", "\\pfnz-srv-028\RetailIMS", FALSE
      If objNet.FolderExists("O:") = True Then
	Log("	   O: drive mapped successfully.")
      Else
	   Log("      Unable to connect O: drive.(" & Err.Number & ", " & Err.Description & ")")
      End If
    End If

	If objNet.FolderExists("U:") = True Then
       Err.Clear
	Log("	   U: drive found.. attempting to disconnect.")
	objNet.FolderExists "U:", TRUE, TRUE
	Do While objNet.FolderExists("U:")
	    If objNet.FolderExists("U:") = False Then
	       Exit Do
	    End If
	    Wscript.Sleep 1000
	    intCounter = intCounter + 1
	    If intCounter = 10 Then
	       Exit Do
	    End If
	Loop
    End If

    If objNet.FolderExists("U:") = False Then
      objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA", FALSE
      If objNet.FolderExists("U:") = True Then
	Log("	   U: drive mapped successfully.")
      Else
	   Log("      Unable to connect U: drive.(" & Err.Number & ", " & Err.Description & ")")
      End If
    End If
	
    'Clr Err Object
    Err.Clear

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_PFNZ_MARKETING

    'Clear Err object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    If objFSO.FolderExists("S:") = True Then
       Err.Clear	  
		Log("	     S: drive found.. attempting to disconnect.")
		objNet.RemoveNetworkDrive "S:", True, True
		Do While objFSO.FolderExists("S:")
			If objFSO.FolderExists("S:") = False Then
			   Exit Do
			End If
			Wscript.Sleep 1000
			intCounter = intCounter + 1
			If intCounter = 10 Then
			   Exit Do
			End If
		Loop
    End If

    If objFSO.FolderExists("S:") = False Then
      objNet.MapNetworkDrive "S:", "\\PFNZ-DAT-002\PFW-Marketing", FALSE
      If objFSO.FolderExists("S:") = True Then
		Log("	     S: drive mapped successfully.")
      Else
		Log("	     Unable to connect S: drive.(" & Err.Number & ", " & Err.Description & ")")
      End If
    End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_PFG_MARKETING

    'Clear Err object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'If objFSO.FolderExists("M:") = True Then
    '   Err.Clear	  
	'	Log("	     M: drive found.. attempting to disconnect.")
	'	objNet.RemoveNetworkDrive "M:", True, True
	'	Do While objFSO.FolderExists("M:")
	'		If objFSO.FolderExists("M:") = False Then
	'		   Exit Do
	'		End If
	'		Wscript.Sleep 1000
	'		intCounter = intCounter + 1
	'		If intCounter = 10 Then
	'		   Exit Do
	'		End If
	'	Loop
    'End If
	
    'If objFSO.FolderExists("P:") = True Then
     '  Err.Clear	  
		'Log("	     P: drive found.. attempting to disconnect.")
	'	objNet.RemoveNetworkDrive "P:", True, True
	'	Do While objFSO.FolderExists("P:")
	'		If objFSO.FolderExists("P:") = False Then
	'		   Exit Do
	'		End If
	'		Wscript.Sleep 1000
	'		intCounter = intCounter + 1
	'		If intCounter = 10 Then
	'		   Exit Do
	'		End If
	'	Loop
    'End If
	
    If objFSO.FolderExists("M:") = False Then
      objNet.MapNetworkDrive "M:", "\\PFG-DAT-001\Marketing", FALSE
      If objFSO.FolderExists("M:") = True Then
		Log("	     M: drive mapped successfully.")
      Else
		Log("	     Unable to connect M: drive.(" & Err.Number & ", " & Err.Description & ")")
      End If
    End If
	
	If objFSO.FolderExists("P:") = False Then
      objNet.MapNetworkDrive "P:", "\\PFG-DAT-001\MarketingBackup", FALSE
      If objFSO.FolderExists("P:") = True Then
		Log("	     P: drive mapped successfully.")
      Else
		Log("	     Unable to connect P: drive.(" & Err.Number & ", " & Err.Description & ")")
      End If
    End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_AS400_Share_Connect

    'Clear Err object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    If objNet.FolderExists("H:") = True Then
       Err.Clear
	Log("	     H: drive found.. attempting to disconnect.")
	objNet.FolderExists "H:", TRUE, TRUE
	Do While objNet.FolderExists("H:")
	    If objNet.FolderExists("H:") = False Then
	       Exit Do
	    End If
	    Wscript.Sleep 1000
	    intCounter = intCounter + 1
	    If intCounter = 10 Then
	       Exit Do
	    End If
	Loop
    End If

    If objNet.FolderExists("H:") = False Then
      objNet.MapNetworkDrive "H:", "\\dealer.powerfarming.co.nz\qdls", FALSE, "qsecofr", "thebrave"
      If objNet.FolderExists("H:") = True Then
	Log("	     H: drive mapped successfully.")
      Else
	   Log("	Unable to connect H: drive.(" & Err.Number & ", " & Err.Description & ")")
      End If
    End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_PFBI

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Map drives
    'objNet.RemoveNetworkDrive "R:"
    'objNet.MapNetworkDrive "R:", "\\SRV-03\IDSE42BI

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_WestCoast

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next
    
    Log("     GROUPJOB_stdgrp_WestCoast")


    'Map drives
    objNet.RemoveNetworkDrive "Q:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "Q:", "\\SRV-06\PFWestCoast$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-PFTG"
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-WCA"
    	
    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\WestCoast.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\WestCoast.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfwestcoast"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "WestCoast Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_AGSOUTHLAND

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next
    
    Log("     GROUPJOB_stdgrp_AGSOUTHLAND")

    ''Map drives
    'objNet.RemoveNetworkDrive "P:"
    'If InStr(strComputerName, "SRV", 1) <> 0 Then
    '  objNet.MapNetworkDrive "P:", "\\SRV-06\PFAGSOUTHLAND$"
    'End If

    'Map drives
    objNet.RemoveNetworkDrive "Q:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "Q:", "\\SRV-06\PFGore$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGGP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGGS"
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGIA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGIP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGIS"
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGIC"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFSIA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFSIS"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFSIP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFSIW"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFSG1"	
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFSG2"	
	objNet.AddWindowsPrinterConnection "\\SRV-06\PFH-RET-PFGA"	
	

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Southland.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Southland.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfsouthland"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Southland Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_AGRILIFE

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_AgriLife")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\AGRILIFE$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGLP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGLA"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\AgriLife.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\AgriLife.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2agrilife"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "AgriLife Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save		

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_MABERTRACTORS

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_MABERTRACTORS")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\MABERTRACTORS$"
    End If

    'Map Printers
	objNet.AddWindowsPrinterConnection "\\SRV-06\MATP"
	objNet.AddWindowsPrinterConnection "\\SRV-06\MATS"
	objNet.AddWindowsPrinterConnection "\\SRV-06\MATA"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Waikato.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Waikato.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfwaikato"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Waikato Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_PFC1

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-01\PFC1"

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_AGEARTH

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next
    
    Log("      GROUPJOB_stdgrp_AGEARTH")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\AGEARTH$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGEA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\AGES"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Whangarei.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Whangarei.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfwhangarei"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Whangarei Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_GISBORNE

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next
    
    Log("     GROUPJOB_stdgrp_GISBORNE")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PFGISBORNE$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\PTGA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PTGB"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Gisborne.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Gisborne.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfgisborne"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Power Farming Gisborne Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save    

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_PFGPowerTurf

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrpPFGPowerTurf")

    'Map drives
    objNet.RemoveNetworkDrive "I:", True, True
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PowerTurfAustralia$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFGMANAGE"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PFGBADMIN"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\PFG PowerTurf.lnk")

	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\PFG PowerTurf.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=PFGPowerTurf"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "PFG PowerTurf Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_HAMILTON

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_HAMILTON")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PFHAMILTON$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\MAHP"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Hamilton.lnk")

	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Hamilton.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfhamilton"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Power Farming Hamilton Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_TRAINING

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    Log("     GROUPJOB_stdgrp_Training")

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Training.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Training.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=F2pftraining"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Famis 2 Training"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_POWERTRAC

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
   On Error Resume Next
    
   Log("     GROUPJOB_stdgrp_POWERTRAC")

    'Map drives
   objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\Powertrac$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\PTHA"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PTHP"
    objNet.AddWindowsPrinterConnection "\\SRV-06\PTHCol2"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Hawke's Bay.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Hawke's Bay.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfhawkesbay"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Hawke's Bay Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_stdgrp_PFROTORUA

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next
    
    Log("     GROUPJOB_stdgrp_PFROTORUA")

    'Map drives
    objNet.RemoveNetworkDrive "P:"
    If InStr(strComputerName, "SRV", 1) <> 0 Then
      objNet.MapNetworkDrive "P:", "\\SRV-06\PFROTORUA$"
    End If

    'Map Printers
    objNet.AddWindowsPrinterConnection "\\SRV-06\SCM1"
    objNet.AddWindowsPrinterConnection "\\SRV-06\SCM2"
    objNet.AddWindowsPrinterConnection "\\SRV-06\SCM3"

    'Create ShortCuts
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Rotorua.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Rotorua.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=srv-06, db=f2pfrotorua"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal\Renault.ico"
	objHelpdeskDesktopLink.Description = "Rotorua Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal"
	objHelpdeskDesktopLink.Save

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GROUPJOB_RETAILADMIN

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next
	
	Log("       Installing Wholesale Printers for Retail Admins.")				
	Call UnMapIDrive
	Call UnMapUDrive
	objNet.MapNetworkDrive "I:", "\\PFNZ-SRV-028\PFWDATA" ,TRUE 				
	objNet.MapNetworkDrive "U:", "\\PFNZ-SRV-028\RETAILDATA", TRUE
	
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\HENGA", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\HENGB", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\HENGD", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\HENGE", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\HENGG", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\HENGH", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\HENGI", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-ASH-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-ASH-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-ASH-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-ASH-UBA", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-AWA-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-AWA-ADM2", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-AWA-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-AWA-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-BAL-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-CAN-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-CAN-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-CAN-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-CAN-SVC2", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-DAR-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-DAR-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-DAR-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-GIS-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-GIS-ADM2", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-GIS-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-GOR-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-GOR-PT2", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-GOR-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-HWK-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-HWK-COL", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-HWK-KON", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-HWK-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-INV-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-INV-ADM2", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-INV-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-INV-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-INV-WRK", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MMM-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MMM-ADM2", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MMM-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MMM-PRT2", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MMM-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MMM-SVC2", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MMM-UGH", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MWT-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MWT-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-MWT-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-NTH-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-NTH-OFF", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-NTH-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-OTA-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-OTA-COL", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-OTA-SAL", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-OTA-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-OTA-URB", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TAR-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TAR-ADM2", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TAR-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TAR-SAL", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TAR-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TIM-ADM", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TIM-COP", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TIM-PRT", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-TIM-SVC", TRUE, TRUE
	objNet.RemovePrinterConnection "\\PFNZ-SRV-028\RET-WST-ADM", TRUE, TRUE
						
	Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-ACC")				
	Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-COL2")  
	Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-COP2")	
	Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-COP2-BW")
	Call MapPrinter("\\PFNZ-SRV-028\PFW-MVL-REC")
	Call MapPrinter("\\PFNZ-SRV-028\PFW-USR-UHM") 
	Call MapPrinter("\\PFNZ-SRV-028\PFW-USR-UKP")
	Call MapPrinter("\\PFNZ-SRV-028\PFW-USR-UKS")
	Call MapPrinter("\\PFNZ-SRV-028\PFW-USR-URE")
	
	Log("       Completed done.")
	
	strDesktop = objWshShell.SpecialFolders("Desktop")
	objFSO.DeleteFile(strDesktop & "\Otago.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Otago.lnk")

	'Create F2 Desktop Shortcuts
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfOtago"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Otago Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save
	
	objFSO.DeleteFile(strDesktop & "\Maber Motors.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Maber Motors.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfmabermotors"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Maber Motors Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save	
	
	objFSO.DeleteFile(strDesktop & "\PF Northland.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\PF Northland.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfwhangarei"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "PF Northland Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save
	
	objFSO.DeleteFile(strDesktop & "\Te Awamutu.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Te Awamutu.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfwaikato"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Te Awamutu Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1,22"
	objHelpdeskDesktopLink.Save	

	objFSO.DeleteFile(strDesktop & "\Timaru.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Timaru.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pftimaru"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Timaru Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save

	objFSO.DeleteFile(strDesktop & "\Gore.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Gore.lnk")
	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfgore"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Gore Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save

	objFSO.DeleteFile(strDesktop & "\Invercargill.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Invercargill.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfsouthland"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Invercargill Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save	
	
		objFSO.DeleteFile(strDesktop & "\Canterbury.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Canterbury.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfcanterbury"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Canterbury Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save	
	
	objFSO.DeleteFile(strDesktop & "\Ashburton.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Ashburton.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfashburton"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Ashburton Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save
	
	objFSO.DeleteFile(strDesktop & "\WestCoast.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\WestCoast.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfwestcoast"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "West Coast Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save	

	objFSO.DeleteFile(strDesktop & "\Taranaki.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Taranaki.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pftaranaki"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Taranaki Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save

	objFSO.DeleteFile(strDesktop & "\Manawatu.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Manawatu.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfmanawatu"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Manawatu Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save	

	objFSO.DeleteFile(strDesktop & "\HawkesBay.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\HawkesBay.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfhawkesbay"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Hawkes Bay Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save
	
	objFSO.DeleteFile(strDesktop & "\Gisborne.lnk")
	Set objHelpdeskDesktopLink = objWshShell.CreateShortcut(strDesktop & "\Gisborne.lnk")

	'Create Helpdesk Desktop Shortcut
	objHelpdeskDesktopLink.TargetPath = "C:\F2ProgramsLocal1.22\FAMIS2000.exe"
	objHelpdeskDesktopLink.Arguments = "svr=pfnz-srv-034, db=f2pfgisborne"
	objHelpdeskDesktopLink.WindowStyle = 1
	objHelpdeskDesktopLink.IconLocation = "%SystemDrive%\F2ProgramsLocal1.22\same.ico"
	objHelpdeskDesktopLink.Description = "Gisborne Famis 2"
	objHelpdeskDesktopLink.WorkingDirectory = "C:\F2ProgramsLocal1.22"
	objHelpdeskDesktopLink.Save
	
    'Disable Error Handling
    On Error Goto 0
	
End Sub
Sub CDSERVER_DRIVE_MAP

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Map drives
	'objNet.RemoveNetworkDrive "J:", True, True
    'objNet.MapNetworkDrive "J:", "\\PFNZ-CDS-001\VOLUMES"
	objNet.RemoveNetworkDrive "K:", True, True
    objNet.MapNetworkDrive "K:", "\\PFNZ-CDS-002\VOLUMES"

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub SetupHelpDesk

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Check OS and run for Win2k systems only.
    If Instr(1, strLocal_OS, "Windows 2000") <> 0 Then

		'Delete old helpdesk icon.
		objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\HelpDesk.lnk")
		objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\HelpDesk.lnk")
		objFSO.DeleteFile("C:\Documents and Settings\All Users\Desktop\IT Training Registration.lnk")
		objFSO.DeleteFile("C:\Documents and Settings\" & strUserName & "\Desktop\IT Training Registration.lnk")

	End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GetComputerName

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'This routine gets the Computer NetBIOS Name of the local system.

    'This procedure has been tested on the following OS's
    '	  1. Windows 2000 Professional
    '	  2. Windows 2000 Server
    '	  3. Windows 98 & SE
    '	  4. Windows 95
    '	  5. Windows NT Workstation 4.0
    '	  6. Windows NT Server 4.0

    'Retrieve Computer NetBIOS Name
    Select Case strLocal_OS

      Case "Windows 95"
	   'Get computer name from registry
	   strComputerName = objWshShell.RegRead(_
	      "HKLM\System\CurrentControlSet\Services\VxD\VNETSUP\ComputerName")

      Case "Windows 98"
	   'Get computer name from registry
	   strComputerName = objWshShell.RegRead(_
	      "HKLM\System\CurrentControlSet\Services\VxD\VNETSUP\ComputerName")

      Case "Windows NT Workstation 4.0"
	   'Get computer name from registry
	   strComputerName = objWshShell.RegRead(_
	      "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")

      Case "Windows NT Server 4.0"
	   'Get computer name from registry
	   strComputerName = objWshShell.RegRead(_
	      "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")

      Case "Windows 2000 Professional"
	   'Get computer name from registry
	   strComputerName = objWshShell.RegRead(_
	      "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")

      Case "Windows 2000 Server"
	   'Get computer name from registry
	   strComputerName = objWshShell.RegRead(_
	      "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")

    End Select	  
	
    'Disable Error Handling
    On Error Goto 0

End Sub


Sub EnforceSettings

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'Enforce Internet Proxy Settings
    objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride", "192.*;168.8.152.101;txtcor02.textronturf.com*;www.anz.com;202.2.59.40;deskbank1.westpac.co.nz;deskbank2.westpac.co.nz;daedong.co.kr;anz.co.nz;coddi.com;savsystem.merlo.com;161.71.70.*;cdserver2;srv-01;srv-02;srv-03;srv-04;srv-05;srv-06;srv-07;srv-08;srv-09;srv-10;srv-11;srv-06;powerlink.powerfarming.co.nz;portal.powerfarming.co.nz;powerlink.pfgaustralia.com.au;intranet.pfgaustralia.com.au<local>"

	'Enforce Newton Settings
	objWshShell.RegWrite "HKCU\Software\Matian.it\CatalX\McCormick\ImagePath", "\\cdserver2\NEWTON43\program files\Newton\drawings\"
	objWshShell.RegWrite "HKCU\Software\Matian.it\CatalX\McCormick\ConnectString", "Provider=SQLOLEDB;Network=dbmssocn;Data Source=srv-03;Initial Catalog=McCormick430;User ID=sa;Password="

    '
    '
    'Enforce Outlook 2000 Settings
    'Outlook message arrival visual notification - ENABLED
    objWshShell.RegWrite "HKCU\Software\Microsoft\Office\9.0\Outlook\Preferences\Notification", 1, "REG_DWORD"

    '
    '
    'Copy Outlook.txt holidays file to local machines
    If objFSO.FileExists("C:\Program Files\Microsoft Office\Office\1033\Outlook.txt") Then
      'The file exists, replace it.
      objFSO.CopyFile "\\powerfarming.co.nz\NETLOGON\Outlook.txt", "C:\Program Files\Microsoft Office\Office\1033\Outlook.txt"
    End If

    '
    '
    'Enable Outlook Administration through Exchange 2000 Server
    'Only implement if running Windows 2000
    'If strLocal_OS = "Windows 2000 Professional" Then
    '   objWshShell.Run "Regedit /S " & "\\powerfarming.co.nz\netlogon\EnableOutlookAdmin.reg", 7, TRUE
    'End If
	
	'
	'
	'Enable Outlook Administration for Outlook 2007	
	objWshShell.RegWrite "HKCU\Software\Policies\Microsoft\Office\12.0\Outlook\Security\AdminSecurityMode", 1, "REG_DWORD"		

    '
    '
    'Enforce Power Policy
    'If strLocal_OS = "Windows 2000 Professional" Then
    '   objWshShell.Run "Regedit /S " & "\\powerfarming.co.nz\netlogon\W2klaptoppwrconf.reg", 7, TRUE
    'End If


	'
	'
	'Enforce IDSe42 Rebranding 122004
	
	'Check log Key		
	strIDSe42_UPDATECODE = objWshShell.RegRead("HKLM\Software\IDS Enterprise Systems Pty Ltd\UpdateCode")		
	If Err.number <> 0 Then
		If Trim(strIDSe42_UPDATECODE) = "" Then
			'Create Key
			objWshShell.RegWrite "HKLM\Software\IDS Enterprise Systems Pty Ltd\UpdateCode", ""
			strIDSe42_UPDATECODE = "0"
		End If				
	End If
	
	If Instr(1,strComputerName, "SRV") = 0 Then
		'strUpdate = objWshShell.Run ("\\srv-01\netlogon\PFWBRAND.hta", 1, TRUE)
	End If
								
    '
    '
    'Enforce IDSe42 access string
    If (objFSO.FileExists("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.ini")) Then
	'Idse42 Software is installed. Check registry entry for mod.
	objWshShell.RegWrite "HKLM\Software\PowerFarming\", ""
	'Ini Vars
	strIDSe42_SYSTEM = ""
	'Attempt to read IDse42 installed & correctly configured Key
	strIDSe42_SYSTEM = objWshShell.RegRead("HKLM\Software\PowerFarming\IDSe42_System")

	'Test for zero length string
	If Len(strIDSe42_SYSTEM) = 0 Then
	   'Entry does not exist - create it
	     'Create RunCount Value
	     objWshShell.RegWrite "HKLM\Software\PowerFarming\IDSe42_System", "SET"
	     
	     'Rename existing file
	     objFSO.CopyFile "C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 " &_
			     "GUI\Settings.ini", "C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.OLD"
	     'DeleteFile
	     objFSO.DeleteFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.ini")
	     'Open renamed file for READING
	     Set txtIDSe42SettingsRD = objFSO.OpenTextFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.OLD", 1)
	     'Open new Settings.INI file for writing
	     Set txtIDSe42SettingsWR = objFSO.OpenTextFile("C:\Program Files\IDS Enterprise Systems Pty Ltd\IDSe42 GUI\Settings.ini", 8, TRUE)

	     'Loop through items in OLD file
	     Do While txtIDSe42SettingsRD.AtEndOfStream <> True
		'ReadLine
		strCurrentLine = txtIDSe42SettingsRD.ReadLine
		'Check for system entry
		If Instr(strCurrentLine, "SYSTEM=") <> 0 Then
		   'Change the value of strCurrentline to new DNS name.
		   strCurrentLine = "SYSTEM=dealer.powerfarming.co.nz"
		    'Write line into new file
		    txtIDSe42SettingsWR.Write strCurrentLine
		    txtIDSe42SettingsWR.WriteLine
		Else
		    'Write line into new file
		    txtIDSe42SettingsWR.Write strCurrentLine
		    txtIDSe42SettingsWR.WriteLine
		End If
	     Loop
	End If
    End If

    'Disable Error Handling
    On Error Goto 0

End Sub

'
'
'...............................................................................................
Sub SetMAPIProfile

    'Clear Err Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next
    
    Exit Sub

    'Check version of Outlook being run
    strOutlookInstPath = objWshShell.RegRead("HKLM\Software\Microsoft\Office\9.0\Outlook\InstallRoot\Path")
    If strOutlookInstPath < 1 Then
	strOutlookInstPath = objWshShell.RegRead("HKLM\Software\Microsoft\Office\10.0\Outlook\InstallRoot\Path")
    End If

    'Set Variables
    strLogonServer = "pfnz-srv-015"

    'Logon NetBIOS LM Name. (UNC Name)
    strSourceDirectory = "\\powerfarming.co.nz\" & "Netlogon"
		   'Server NETLOGON share path.

    'Set the default MAPI profile
    If objWshShell.RegRead("HKCU\Software\Microsoft\Windows Messaging Subsystem\Profiles\DefaultProfile") <> strUserName Then
       'Write the default MAPI profile
       objWshShell.RegWrite "HKCU\Software\Microsoft\Windows Messaging Subsystem\Profiles\DefaultProfile", strUserName
    End If

    If objWshShell.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\DefaultProfile") <> strUserName Then
       'Write the default MAPI profile
       objWshShell.RegWrite "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\DefaultProfile", strUserName
    End If

    'This routine requires the presence of both the FIXPRF.EXE and NEWPROF.EXE in the
    'source (usually NETLOGON directory).

    'Copy MapiSvc.INF file to local system directory for Exchange entries
    objFSO.CopyFile strSourceDirectory & "\" & "MAPISVC.INF", strSystemRoot & "\MAPISVC.INF"

    'Copy OutlBar.inf file to the installed Outlook directory
    '	  By default this is at C:\Program Files\Microsoft Office\Office\1033
    '	  This file can be customized to setup default Outlook Shortcuts in Outlook.
    objFSO.CopyFile "\\powerfarming.co.nz\NETLOGON\OUTLBAR.INF", "C:\PROGRAM FILES\MICROSOFT OFFICE\OFFICE\1033\OUTLBAR.INF"

    'Copy source PRF file to local system (assumes C: is a local & writable drive)
    objFSO.CopyFile strSourceDirectory & "\" & "OUTLPROF.PRF", "C:\"


    'If InStr(strOutlookInstPath, "OFFICE11") = 0 Then
    '  'Run NEWPROF only if version is NOT 10.0 or 11.0
    '  If InStr(strOutlookInstPath, "Office10") = 0 Then
    '	'Modify template PRF to be user specific
    '	objWshShell.Run strSourceDirectory & "\" & "FIXPRF C:\OUTLPROF.PRF " & strUserName & " " & strUserName & " " & strLogonServer, 7, TRUE
    '	'Create profile based on PRF file
    '	objWshShell.Run strSourceDirectory & "\" & "NEWPROF -P " & "C:\OUTLPROF.PRF" & " -X", 7, TRUE
    '  End If
    'Else
    '  'Set the default MAPI profile
    '  If objWshShell.RegRead("HKCU\Software\Microsoft\Windows Messaging Subsystem\Profiles\DefaultProfile") <> "PFAUTOPROFILE" Then
    '	 'Write the default MAPI profile
    '	 objWshShell.RegWrite "HKCU\Software\Microsoft\Windows Messaging Subsystem\Profiles\DefaultProfile", "PFAUTOPROFILE"
    ' End If

    '  If objWshShell.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\DefaultProfile") <> "PFAUTOPROFILE" Then
	 'Write the default MAPI profile
    '	 objWshShell.RegWrite "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\DefaultProfile", "PFAUTOPROFILE"
    '  End If
    'End If

    'Disable Error Handling
    On Error Goto 0

End Sub

'
'
'...............................................................................................
Sub GetUserName

    'Resume execution even if errors occur
    On Error Resume Next

    'This routine takes into account the fact that at logon, the UserName function is not
    'immediately available. Therefore, it loops (maximum 100000 times) until the variable has
    'been filled with some (any) text.

    Do While strUserName = ""

       'Setup UserName var
       strUserName = objNet.UserName

      'Setup loop threshold counter
      Counter = Counter + 1

      'Shutdown script after more than 10000 loops
      If Counter > 100000 Then
	 Wscript.Quit
      End If

    Loop

    'Disable Error Handling
    On Error Goto 0

End Sub
'...............................................................................................


Sub Local_OS()

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'This procedure will report whether the local OS is one of the following:
    '	  1. Windows 95
    '	  2. Windows 98
    '	  3. Windows NT Workstation 4.0
    '	  4. Windows NT Server 4.0
    '	  5. Windows 2000 Professional
    '	  6. Windows 2000 Server

    'This procedure has been tested on the following OS's
    '	  1. Windows 2000 Professional
    '	  2. Windows 2000 Server
    '	  3. Windows 98 & SE
    '	  4. Windows 95
    '	  5. Windows NT Workstation 4.0
    '	  6. Windows NT Server 4.0 (TEST PENDING)

    'Check to see whether local system is NT or 9x
    'Attempt to access NT registry key
    objWshShell.RegRead ("HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType")

    'Run test on error number
    If Err.Number <> -2147024894 Then
       'Local OS must be NT based
       strLocal_OS = objWshShell.RegRead(_
	  "HKLM\Software\Microsoft\Windows NT\CurrentVersion\ProductName")	  
       'Test for error state
       If Err.Number = 0 Then
	 'Test for Server or Workstation based OS
	 If objWshShell.RegRead(_
	    "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType") = "WinNT" Then
	    'OS is 2k based Workstation
	    strLocal_OS = "Windows 2000 Professional"
	 ElseIf objWshShell.RegRead(_
	    "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType") = "LanmanNT" OR _
	    objWshShell.RegRead(_
	    "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType") = "ServerNT" Then
	    'OS is 2k based Server
	    strLocal_OS = "Windows 2000 Server"
	 End If
       Else
	 'Test for Server or Workstation based OS
	 If objWshShell.RegRead(_
	    "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType") = "WinNT" Then
	    'OS is NT4.0 based Workstation
	    strLocal_OS = "Windows NT Workstation 4.0"
	 ElseIf objWshShell.RegRead(_
	    "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType") = "LanmanNT" Then
	    'OS is NT4.0 based Server
	    strLocal_OS = "Windows NT Server 4.0"
	 End If
       End If
    ElseIf Err.Number = -2147024894 Then
       'Local OS must be legacy based
       strLocal_OS = objWshShell.RegRead(_
	  "HKLM\Software\Microsoft\Windows\CurrentVersion\ProductName")
       'Remove Microsoft OS Prefix if it exists
       If Instr(strLocal_OS, "Microsoft") <> 0 OR Instr(strLocal_OS, "Microsoft") <> NULL Then
	  'Strip the "Microsoft" bit off (to standardise with rest of script)
	  strLocal_OS = Mid(strLocal_OS, 11)
       End If

    End If

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GetLocalSystemRoot

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'This script returns the root directory for the OS ie. C:\Windows OR C:\WINNT

    'This procedure has been tested on the following OS's
    '	  1. Windows 2000 Professional
    '	  2. Windows 2000 Server
    '	  3. Windows 98 & SE
    '	  4. Windows 95
    '	  5. Windows NT Workstation 4.0
    '	  6. Windows NT Server 4.0 (TEST PENDING)

    'Retrieve Systemroot from registry
    Select Case strLocal_OS

      Case "Windows 95"
	       'Setup SystemRoot
	       strSystemRoot = objWshShell.RegRead(_
		  "HKLM\Software\Microsoft\Windows\CurrentVersion\SystemRoot")
      Case "Windows 98"
	       'Setup SystemRoot
	       strSystemRoot = objWshShell.RegRead(_
		  "HKLM\Software\Microsoft\Windows\CurrentVersion\SystemRoot")
      Case "Windows NT Workstation 4.0"
	       'Setup SystemRoot
	       strSystemRoot = objWshShell.RegRead(_
		  "HKLM\Software\Microsoft\Windows NT\CurrentVersion\SystemRoot")
      Case "Windows NT Server 4.0"
	       'Setup SystemRoot
	       strSystemRoot = objWshShell.RegRead(_
		  "HKLM\Software\Microsoft\Windows NT\CurrentVersion\SystemRoot")
      Case "Windows 2000 Professional"
	       'Setup SystemRoot
	       strSystemRoot = objWshShell.RegRead(_
		  "HKLM\Software\Microsoft\Windows NT\CurrentVersion\SystemRoot")
      Case "Windows 2000 Server"
	       'Setup SystemRoot
	       strSystemRoot = objWshShell.RegRead(_
		  "HKLM\Software\Microsoft\Windows NT\CurrentVersion\SystemRoot")
    End Select

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub GetIEVersion

    'Clear Error Object
    Err.Clear

    'Enable Error Handling
    On Error Resume Next

    'This routine gets the Internet Explorer version on the local system.

    'This procedure has been tested on the following OS's
    '	  1. Windows 2000 Professional
    '	  2. Windows 2000 Server
    '	  3. Windows 98 & SE
    '	  4. Windows 95
    '	  5. Windows NT Workstation 4.0
    '	  6. Windows NT Server 4.0 (TEST PENDING)

    'Retrieve IE Version Number from registry
    Select Case strLocal_OS

      Case "Windows 95"
	   'Setup IE Version
	   strIEVersion = objWshShell.RegRead(_
	      "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")

      Case "Windows 98"
	   'Setup IE Version
	   strIEVersion = objWshShell.RegRead(_
	      "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")

      Case "Windows NT Workstation 4.0"
	   'Setup IE Version
	   strIEVersion = objWshShell.RegRead(_
	      "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")

      Case "Windows NT Server 4.0"
	   'Setup IE Version
	   strIEVersion = objWshShell.RegRead(_
	      "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")

      Case "Windows 2000 Professional"
	   'Setup IE Version
	   strIEVersion = objWshShell.RegRead(_
	      "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")

      Case "Windows 2000 Server"
	   'Setup IE Version
	   strIEVersion = objWshShell.RegRead(_
	      "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")

    End Select

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub Log(strLog)
	Err.Clear
	On Error Resume Next	

	If objFSO.FolderExists("c:\Support") <> True Then	
		objFSO.CreateFolder "c:\Support"
	End If
	
	'Open / Create Text File
	Set fileLogon = objFSO.OpenTextFile("C:\Support\LogonScriptLog.txt", 8, True)

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

Sub Mailer

    'Clear Error Object
    Err.Clear

    'Enable error handling
    On Error Resume Next

    'Create Object
    Set objSMTPMail = CreateObject("SMTPControl.SMTP")

    'Setup SMTPMail Parameters
    objSMTPMail.Server = strSMTPServer
    objSMTPMail.MailFrom = strSMTPMailFrom
    objSMTPMail.SendTo = strSMTPSendTo
    objSMTPMail.MessageSubject = strSMTPMessageSubject
    objSMTPMail.MessageText = strSMTPMessageText
    objSMTPMail.Connect

    'Test Mail Send for Errors
    If objSMTPMail.Status = "SMTP control error" Then
       'DEBUGGING
       'Wscript.Echo "There was an error sending the email!"
    End If

    'Re-Init SMTP Mailer Vars
    strSMTPServer = ""
    strSMTPMailFrom = ""
    strSMTPSendTo = ""
    strSMTPMessageSubject = ""
    strSMTPMessageText = ""

    'Kill Object
    Set objSMTPMail = Nothing

    'Disable Error Handling
    On Error Goto 0

End Sub

Sub MapPrinter(printPath)

	'Clear Err Object
	Err.Clear 
	
	'Enable Error Handling
	On Error Resume Next
	
	Const HKCU = &H80000001
	sKey = "Software\Microsoft\Windows NT\CurrentVersion\Devices"
	Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	iSuccess = oReg.GetStringValue(HKCU, sKey, printPath, sData)	
	
	If iSuccess = 1 Then
		Log("        Mapped New printer: " & UCASE(printPath))
		objNet.AddWindowsPrinterConnection printPath
	Else
		Log("        Found Existing Printer: " & UCASE(printPath))	
	End If
	
	'64bit Print Server v10
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t100", "\\PFNZ-SRV-028\PFW-USR-UDB", "REG_SZ" 'Donna Black
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t101", "\\PFNZ-SRV-028\RET-ASH-PRT", "REG_SZ" 'Ashburton Parts
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t102", "\\PFNZ-SRV-028\RET-INV-WRK", "REG_SZ" 'Invercargill Workshop
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t103", "\\PFNZ-SRV-028\PFW-MVL-WAR", "REG_SZ" 'Morrinsville Warranty
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t104", "\\PFNZ-SRV-028\RET-TIM-PRT", "REG_SZ" 'Timaru Parts
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t105", "\\PFNZ-SRV-028\PFH-RET-WCA", "REG_SZ" 'WestCoast Admin
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t106", "\\PFNZ-SRV-028\PFW-USR-URE", "REG_SZ" 'Morrinsvile Raewyn Evans
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t107", "\\PFNZ-SRV-028\PFW-USR-UKA", "REG_SZ" 'Morrinsville Karin Adams
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t108", "\\PFNZ-SRV-028\PFW-USR-MGT", "REG_SZ" 'Morrinsville Management ( Shelley Burger )
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t109", "\\PFNZ-SRV-028\PFW-USR-ASS", "REG_SZ" 'Morrinsville Assembly
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t110", "\\PFNZ-SRV-028\PFW-USR-UKP", "REG_SZ" 'Morrinsvile Karen Parton	

	'32bit Print Server v10
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t200", "\\PFNZ-SRV-029\PFW-USR-UDB", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t201", "\\PFNZ-SRV-029\RET-ASH-PRT", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t202", "\\PFNZ-SRV-029\RET-INV-WRK", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t203", "\\PFNZ-SRV-029\PFW-MVL-WAR", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t204", "\\PFNZ-SRV-029\RET-TIM-PRT", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t205", "\\PFNZ-SRV-029\PFH-RET-WCA", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t206", "\\PFNZ-SRV-029\PFW-USR-URE", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t207", "\\PFNZ-SRV-029\PFW-USR-UKA", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t208", "\\PFNZ-SRV-029\PFW-USR-MGT", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t209", "\\PFNZ-SRV-029\PFW-USR-ASS", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\10.0\General\cPrintAsImage\t210", "\\PFNZ-SRV-029\PFW-USR-UKP", "REG_SZ" 'Morrinsvile Karen Parton		
	'64bit Print Server v11
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t100", "\\PFNZ-SRV-028\PFW-USR-UDB", "REG_SZ" 'Donna Black
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t101", "\\PFNZ-SRV-028\RET-ASH-PRT", "REG_SZ" 'Ashburton Parts
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t102", "\\PFNZ-SRV-028\RET-INV-WRK", "REG_SZ" 'Invercargill Workshop
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t103", "\\PFNZ-SRV-028\PFW-MVL-WAR", "REG_SZ" 'Morrinsville Warranty
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t104", "\\PFNZ-SRV-028\RET-TIM-PRT", "REG_SZ" 'Timaru Parts
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t105", "\\PFNZ-SRV-028\PFH-RET-WCA", "REG_SZ" 'WestCoast Admin
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t106", "\\PFNZ-SRV-028\PFW-USR-URE", "REG_SZ" 'Morrinsvile Raewyn Evans
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t107", "\\PFNZ-SRV-028\PFW-USR-UKA", "REG_SZ" 'Morrinsville Karin Adams
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t108", "\\PFNZ-SRV-028\PFW-USR-MGT", "REG_SZ" 'Morrinsville Management ( Shelley Burger )
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t209", "\\PFNZ-SRV-028\PFW-USR-ASS", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t110", "\\PFNZ-SRV-028\PFW-USR-UKP", "REG_SZ" 'Morrinsvile Karen Parton		
	
	'32bit Print Server v11
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t200", "\\PFNZ-SRV-029\PFW-USR-UDB", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t201", "\\PFNZ-SRV-029\RET-ASH-PRT", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t202", "\\PFNZ-SRV-029\RET-INV-WRK", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t203", "\\PFNZ-SRV-029\PFW-MVL-WAR", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t204", "\\PFNZ-SRV-029\RET-TIM-PRT", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t205", "\\PFNZ-SRV-029\PFH-RET-WCA", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t206", "\\PFNZ-SRV-029\PFW-USR-URE", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t207", "\\PFNZ-SRV-029\PFW-USR-UKA", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t208", "\\PFNZ-SRV-029\PFW-USR-MGT", "REG_SZ"	
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t209", "\\PFNZ-SRV-029\PFW-USR-ASS", "REG_SZ"
	objWshShell.RegWrite "HKCU\SOFTWARE\Adobe\Acrobat Reader\11.0\General\cPrintAsImage\t210", "\\PFNZ-SRV-029\PFW-USR-UKP", "REG_SZ" 'Morrinsvile Karen Parton	
		
	'Disable Error Handling
	On Error Goto 0	
	
End Sub

Function CTimeStamp(dt)

	'Clear Err Object
	Err.Clear 
	
	'Enable Error Handling
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
	
	'Disable Error Handling
	On Error Goto 0

End Function

Sub GetCurrentTimeStamp

	'Clear Err Object
	Err.Clear 
	
	'Enable Error Handling
	On Error Resume Next

	'Set Dates/Times
	strMonth = Month(Now)
	strDay = Day(Now)
	strHour = Hour(Now)
	strMin = Minute(Now)		
	
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
	
	'Current Time Stamp
	strCurrentDateTimeStamp = Year(Now) & strMonth & strDay & strHour & strMin	
	
	'Disable Error Handling
	On Error Goto 0

End Sub

Sub EnableOutlookCachedExchangeModeAndPublicFolderFavoritesForLaptops

	On Error Resume Next

	If InStr(strComputername, "NPFG") > 0 Or _
		InStr(strComputername, "NPFW") > 0 Or _
			InStr(strComputername, "NHAU") > 0 Or _
				InStr(strComputername, "UPFW") > 0 Or _
					InStr(strComputername, "UPFG") > 0 Then


		'Outlook 2010 
		DefaultProfile = ""
		DefaultProfile = Trim(objWshShell.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\DefaultProfile"))

		If Trim(DefaultProfile) <> "" Then
			DefaultProfile = Trim(objWshShell.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\DefaultProfile"))
			Version = "2010"
			Log("EnableOutlookCachedExchangeModeAndPublicFolderFavoritesForLaptopsPath: Outlook 2010 Detected")	 
		Else
			'2013?
			Log("EnableOutlookCachedExchangeModeAndPublicFolderFavoritesForLaptopsPath: Outlook 2013 Detected Maybe?")	 			
			DefaultProfile = Trim(objWshShell.RegRead("HKCU\Software\Microsoft\Office\15.0\Outlook\DefaultProfile"))
			Version = "2013"
		End If
					
		If Trim(DefaultProfile) = "" Then
			Exit Sub
		End If					
				
		Const HKEY_CURRENT_USER = &H80000001	
		Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")	
		If Version = "2010" Then
			strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\" & DefaultProfile
		Else
			strKeyPath = "Software\Microsoft\Office\15.0\Outlook\Profiles\" & DefaultProfile
		End If		
		objReg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrSubKeys
		
		For Each Subkey in arrSubKeys		
			strValueName = "00036601"
			objReg.GetBinaryValue HKEY_CURRENT_USER,strKeyPath + "\" + Subkey,strValueName,strValue		
			If Not IsNull(strValue) Then		
				'arrValues = Array(132,25,0,0)  ' enable caching
				arrValues = Array(132,5,0,0)  ' enable caching with public Folders
				'arrValues = Array(4,16,0,0)   ' disable caching		
				errReturn = objReg.SetBinaryValue (HKEY_CURRENT_USER, strKeyPath + "\" + Subkey, strValueName, arrValues)		
			End If	
		Next
	
	End If

End Sub

Sub DefaultPrinterBalloonNotify
	On Error Resume Next
	If objFSO.FileExists("C:\Support\Printer-Ink.ico") = False Then
		objFSO.CopyFile "\\powerfarming.co.nz\netlogon\svn-netlogon\login\Printer-Ink.ico", "C:\Support\"
	End If
	
	retval = objWshShell.Run ("powershell &'" & "\\powerfarming.co.nz\NETLOGON\svn-netlogon\Login\BalloonNotifyDefaultPrinter.ps1" & "'", 0, FALSE)
End Sub

Sub CleanUp

	'Clear Err Object
	Err.Clear

	'Enable Error Handling
	On Error Resume Next

		'Set default printer to whatever it was set to at the start of the script in case
		'AddWindowsPrinterConnection has changed it.
		If Trim(strDefaultPrinterPath) <> "" And Trim(strDefaultPrinterPath) <> "\" Then
			objNET.SetDefaultPrinter strDefaultPrinterPath
		End If
		
		Call DefaultPrinterBalloonNotify
	
	sTotalElapsed = DateDiff("s", dtStartNow, Now)

	 'Logging
	 Log("Logon completed successfully. (Total Time: " & sTotalElapsed & "secs)")
	 Log("")
	 Log("*************************************************************************************************")
	 Log("")

	'Disable Error Handling
	On Error Goto 0

End Sub

Function RedirectMyDoc(path)
  Err.Clear
  On Error Resume Next
  If objFSO.FolderExists(path) = True Then
  objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal", path, "REG_SZ"
		If Trim(strComputerName) = "HAU-SRV-004"  Then
			 objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favourites", path, "REG_SZ"
			  objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal", path, "REG_SZ"
			   objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favourites", path, "REG_SZ"
		End If
     Log("My Documents rerouted to " & path)
  Else
     Log("My Documents reroute FAILED to " & path)
  End If
  On Error Goto 0
End Function

Function RedirectMyDoc(path)
  Err.Clear
  On Error Resume Next
  If objFSO.FolderExists(path) = True Then
  objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Personal", path, "REG_SZ"
		If Trim(strComputerName) = "HAU-SRV-004"  Then
			 objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favourites", path, "REG_SZ"
			  objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal", path, "REG_SZ"
			   objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favourites", path, "REG_SZ"
		End If
     Log("My Documents rerouted to " & path)
  Else
     Log("My Documents reroute FAILED to " & path)
  End If
  On Error Goto 0
End Function

Function RedirectMyMedia(path)
  Err.Clear
  On Error Resume Next
  If objFSO.FolderExists(path) = True Then

		If Trim(strComputerName) = "HAU-SRV-004"  Then
		  searchStr="Music"
		  If instr(path, searchStr)<>0 Then
			 objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Music", path, "REG_SZ"
			  objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Music", path, "REG_SZ"
		  End If
		  
		  searchStr="Pictures"
		  If instr(path, searchStr)<>0 Then
			 objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\My Pictures", path, "REG_SZ"
			  objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures", path, "REG_SZ"
		  End If
		 End If
     Log("My Media folders rerouted")
  Else
     Log("My Media folders reroute FAILED")
  End If
  On Error Goto 0
End Function

Function CreateHomeDir(strUser, strHome)
  On Error Resume Next
  Log("Home Folder.")
  CreateHomeDir = True
  Set objShell = CreateObject("Wscript.Shell")
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  'If objFileExists(strSystemRoot & "\System32\cacls.exe") = False Then
  '   Log("Cacls not found.")
  '   objFSO.CopyFile strLogonServer & "\Netlogon\CommonApps\cacls.exe", strSystemRoot & "\System32\cacls.exe"
  '   If objFileExists(strSystemRoot & "\System32\cacls.exe") = False Then
  '	 CreateHomeDir = False
  '	 Wscript.Quit
  '   End If
  'End If
  'Wscript.Echo "ASS!"
  If Right(Trim(strHome), 1) <> "\" Then
     strHome = Trim(strHome) & "\"
  End If
  strHomeFolder = strHome & strUser
  Log(strHomeFolder)
  If strHomeFolder <> "" Then
    Log("Creating Home folder for user " & strUser & " at " & strHome & ".")
     If Not objFSO.FolderExists(strHomeFolder) Then
     On Error Resume Next
       objFSO.CreateFolder strHomeFolder
       If Err.Number <> 0 Then	     
	CreateHomeDir = False
	Log("Folder was not created.")
       End If
       On Error GoTo 0
     End If
     If objFSO.FolderExists(strHomeFolder) Then
       ' Assign user permission to home folder.
       intRunError = objShell.Run("%COMSPEC% /c Echo Y| cacls "_
       & strHomeFolder & " /t /c /g Administrators:f "_
       & strUser & ":F", 2, True)
	  If intRunError <> 0 Then
	   CreateHomeDir = False
	   Log("Cacls command failed.")
	  End If
     End If
  End If

  If Trim(strComputerName) = "SRV-11" OR Trim(strComputerName) = "SRV-09" OR Trim(strComputerName) = "HAU-SRV-004" OR Trim(strComputerName) = "PFNZ-SRV-032" Then
     RedirectMyDoc(strHomeFolder)
  End If

End Function


Function CreateMediaDir(strUser, strHome)
  On Error Resume Next
  Log("Media Folders.")
  CreateMediaDir = True
  Set objShell = CreateObject("Wscript.Shell")
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  'If objFileExists(strSystemRoot & "\System32\cacls.exe") = False Then
  '   Log("Cacls not found.")
  '   objFSO.CopyFile strLogonServer & "\Netlogon\CommonApps\cacls.exe", strSystemRoot & "\System32\cacls.exe"
  '   If objFileExists(strSystemRoot & "\System32\cacls.exe") = False Then
  '	 CreateHomeDir = False
  '	 Wscript.Quit
  '   End If
  'End If
  'Wscript.Echo "ASS!"
  If Right(Trim(strHome), 1) <> "\" Then
     strHome = Trim(strHome) & "\"
  End If
  strMusicFolder = strHome & strUser & "\Music"
  strPicsFolder = strHome & strUser & "\Pictures"
  Log(strMusicFolder)
  Log(strPicsFolder)
  If strMusicFolder <> "" Then
    Log("Creating Music folder for user " & strUser & " at " & strHome & ".")
     If Not objFSO.FolderExists(strMusicFolder) Then
     On Error Resume Next
       objFSO.CreateFolder strMusicFolder
       If Err.Number <> 0 Then	     
	CreateMediaDir = False
	Log("Folder was not created.")
       End If
       On Error GoTo 0
     End If
     If Trim(strComputerName) = "HAU-SRV-004" Then
     RedirectMyMedia(strMusicFolder)
	 End If
  End If
  
   If strPicsFolder <> "" Then
    Log("Creating Pictures folder for user " & strUser & " at " & strHome & ".")
     If Not objFSO.FolderExists(strPicsFolder) Then
     On Error Resume Next
       objFSO.CreateFolder strPicsFolder
       If Err.Number <> 0 Then	     
	CreateMediaDir = False
	Log("Folder was not created.")
       End If
       On Error GoTo 0
     End If
     If Trim(strComputerName) = "HAU-SRV-004" Then
     RedirectMyMedia(strPicsFolder)
	 End If
  End If

End Function
'***************************************** DISCLAIMERTOOL DEPOYMENT *************************************************

Function NetFrameworkInstall(InstallMsi)

	 'Clr Err Object
	 Err.Clear

	 'Enable Error Handling
	 On Error Resume Next

	 If MSISoftwareInstalled("framework") = False Then
	  Log("The .NET Framework was not found. Installing.")
	  Dim msiObject, msiProduct, strProdList, strProdInfo, msiProdVersion
	  Set msiObject = Wscript.CreateObject("WindowsInstaller.Installer")
	  msiObject.UILevel = 3 + 64
	  msiObject.InstallProduct InstallMsi
	   If MSISoftwareInstalled("framework") = True Then
	      NetFrameworkInstall = True
	   Else
	      NetFrameworkInstall = False
	   End If
	End If

	 'Disable Error Hanlding
	 On Error Goto 0

End Function

Function MSISoftwareInstalled(NameSearch)

  On Error Resume Next

  Const wbemFlagReturnImmediately = &h10
  Const wbemFlagForwardOnly = &h20
    MSISoftwareInstalled = FALSE

  arrComputers = Array(".")
  For Each strComputer In arrComputers
     Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
     Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_SoftwareFeature", "WQL", _
					    wbemFlagReturnImmediately + wbemFlagForwardOnly)
     For Each objItem In colItems
		  'Wscript.Echo UCase(objItem.ProductName) & " " & UCase(NameSearch)
	 If InStr(UCase(objItem.ProductName), UCase(NameSearch)) <> 0 Then
	      MSISoftwareInstalled = TRUE
	      Exit For
	 End If
     Next
  Next

End Function


'***************************************** MAILARCHIVE FOLDER DEPOYMENT *******************************************

Sub AddOutlookMailArchiveFolder
  Err.Clear
  On Error Resume Next
  'If InStr(strComputerName, "NPFG") <> 0 Or InStr(strComputerName, "PPFG") <> 0 Then
    Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
    tfOutlookActive = FALSE
    intAddActivitesOutlookMod = 0
    Do While intAddActivitesOutlookMod < 100
       Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)
       For Each objItem in colItems
	   If objItem.Description = "OUTLOOK.EXE" Then
	      tfOutlookActive = TRUE
	      intAddActivitesOutlookMod = 100
	   End If
       Next
       Wscript.Sleep 3000
       intAddActivitesOutlookMod = intAddActivitesOutlookMod + 1
    Loop
  
    If tfOutlookActive = TRUE Then
       Wscript.Sleep 10000
			On Error Resume Next
			Const olFolderInbox = 6			
			Set objOutlook = CreateObject("Outlook.Application")
			Set objNamespace = objOutlook.GetNamespace("MAPI")
			Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)			
			strFolderName = objFolder.Parent
			Set objMailbox = objNamespace.Folders(strFolderName)			
			Set objNewFolder = objMailbox.Folders.Add("Mail Archive")
			objNewFolder.WebViewURL = "http://mailarchive.powerfarming.co.nz/outlook.do"
			objNewFolder.WebViewOn = True       
    End If
  On Error Goto 0
End Sub

'***************************************** PFG OUTLOOK DEPOYMENT *************************************************

Sub AddFav_PFG_Australia_Contacts
  Err.Clear
  On Error Resume Next
  If InStr(strComputerName, "NPFG") <> 0 Or InStr(strComputerName, "PPFG") <> 0 Then
    Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
    tfOutlookActive = FALSE
    intAddActivitesOutlookMod = 0
    Do While intAddActivitesOutlookMod < 100
       Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)
       For Each objItem in colItems
	   If objItem.Description = "OUTLOOK.EXE" Then
	      tfOutlookActive = TRUE
	      intAddActivitesOutlookMod = 100
	   End If
       Next
       Wscript.Sleep 3000
       intAddActivitesOutlookMod = intAddActivitesOutlookMod + 1
    Loop
  
    If tfOutlookActive = TRUE Then
       Wscript.Sleep 10000
       If CheckFavExists("PFG Australia Contacts") = False Then
	 Set objOutlook = CreateObject("Outlook.Application")
	 Set objOutlookNameSpace = objOutlook.GetNameSpace("MAPI")
	 Set objOutlookPublicFolders = objOutlookNameSpace.Folders("Public Folders")
	 Set objOutlookAllPublicFolders = objOutlookPublicFolders.Folders("All Public Folders")
	 Set objOutlookAllPublicFolders_PFG_AUSTRALIA = objOutlookAllPublicFolders.Folders("PFG Australia")
	 Set objOutlookAllPublicFolders_PFG_AUSTRALIA_PFG_SHIPPING = objOutlookAllPublicFolders_PFG_AUSTRALIA.Folders("PFG Australia Contacts")
	 objOutlookAllPublicFolders_PFG_AUSTRALIA_PFG_SHIPPING.AddToPFFavorites
	 Wscript.Sleep 10000
	 Call AddOAB_PFG_Australia_Contacts
       End If
    End If
  End If
  On Error Goto 0
End Sub

Sub AddOAB_PFG_Australia_Contacts
   Err.Clear
   On Error Resume Next
   Set objOutlook = CreateObject("Outlook.Application")
   Set objOutlookNameSpace = objOutlook.GetNameSpace("MAPI")
   Set objOutlookPublicFolders = objOutlookNameSpace.Folders("Public Folders")
   Set objOutlookFavorites = objOutlookPublicFolders.Folders("Favorites")
   Set objOutlookPFGAustraliaContacts = objOutlookFavorites.Folders("PFG Australia Contacts")
   objOutlookPFGAustraliaContacts.ShowAsOutlookAB = True
   On Error Goto 0
End Sub

Function CheckFavExists(favname)
   Err.Clear
   On Error Resume Next
   CheckFavExists = False
   Set objOutlook = CreateObject("Outlook.Application")
   Set objOutlookNameSpace = objOutlook.GetNameSpace("MAPI")
   Set objOutlookPublicFolders = objOutlookNameSpace.Folders("Public Folders")
   Set objOutlookFavorites = objOutlookPublicFolders.Folders("Favorites").Folders
   For Each itm In objOutlookFavorites
     If Trim(favname) = Trim(itm.Name) Then
	CheckFavExists = True
     End If
   Next
   On Error Goto 0
End Function

Sub KillImpAdmin
	Err.Clear
	On Error Resume Next

	strComputer = "."
	Set colProcesses = GetObject("winmgmts:" & _
	   "{impersonationLevel=impersonate}!\\" & strComputer & _
	   "\root\cimv2").ExecQuery("Select * from Win32_Process")

	'Init. vars
	tfCont = False


	    For Each objProcess in colProcesses

		Return = objProcess.GetOwner(strNameOfUser)
		If Return <> 0 Then
		    'Wscript.Echo "Could not get owner info for process " & _
		    '	 objProcess.Name & VBNewLine _
		    '	 & "Error = " & Return
		Else
		    'Wscript.Echo strNameOfUser & " " & objProcess.Description
		    If UCase(strNameOfUser) = UCase("adminjobuser") And objProcess.Description = "ImpAdmin.exe" Then
		       objProcess.Terminate
		    End If
		    'Wscript.Echo "Process " _
		    '	 & objProcess.Name & " is owned by " _
		    '	 & "\" & strNameOfUser & "."
		End If
	    Next

	    ''Loop through processes on the workstation and Kill and instances for ImpAdmin.exe
	    'Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)
	    'For Each objItem in colItems
	    '
	    '		 Wscript.Echo objItem.Description & " is owned by " & objItem.GetOwner(strUser)
	    '
	    '		 'Check for Administrator user only
	    '		 If UCase(objItem.GetOwner(strUser)) = UCase("/Administrator") Then
	    '		   'Stop Impromptu if it is found to be running
	    '		   If objItem.Description = "ImpAdmin.exe" Then
	    '			   'Kill Process
	    '			   objItem.Terminate
	    '			   Call LogEvent("PROC:KILLIMPADMIN", "NOTIFY", "Killed Instance of Impromptu.")
	    '		   End If
	    '		 End If
	    '	 Next
	    '	 'Kill colItems Object
	    '	 Set colItems = Nothing
	    '
		'Init.
		tfCont = True
	    'Loop through processes on the workstation and check for instances of ImpAdmin
	    'Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)
	    'For Each objItem in colItems
	    '		 'Check for running instances for ImpAdmin.exe
	    '		 If objItem.Description = "ImpAdmin.exe" Then
	    '			 'Keep looping as ImpAdmin was still there
	    '			 tfCont = False
	    '		 End If
	    '	 Next
	    '	 'Kill colItems Object
	    '	 Set colItems = Nothing
	    '
	'Loop

	'Disable Error Handling
	On Error Goto 0

End Sub

