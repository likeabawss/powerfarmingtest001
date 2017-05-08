
Sub PartsCatMain

    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")

	EventLog("Parts Cat Script Starts")
	
	Call GetUserName
	Call PartsCat_SetupLocalResources
	Call DisableScreenSaver
	
	EventLog("Parts Cat Script Ends")
	
End Sub

Function GetCustomerType()

    On Error Resume Next

    Set cn = CreateObject("ADODB.Connection")
    cn.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-028" & ";Initial Catalog=" & "LauncherV2" & ";Integrated Security=SSPI;"

    Set rs = CreateObject("ADODB.Recordset")
    rs.ActiveConnection = cn
    sql = "SELECT [LauncherV2].[dbo].[GetCustomerType] ('pfw','" & strUserName & "')"	
    rs.Open sql,cn,1,1
    rs.MoveFirst

    GetCustomerType = rs(0)

	If err.number <> 0 then
		EventLog("GetCustomerType|" & err.number & "|" & err.description)
	End If
	
    On Error Goto 0
    err.clear

End Function

Sub DisableScreenSaver

	Wscript.Interactive = True	
	On Error Resume Next
	
	objWshShell.RegWrite "HKCU\Control Panel\Desktop\ScreenSaveActive", "1", "REG_SZ"			
	
	Err.Clear

End Sub

Sub PartsCat_SetupLocalResources

	Wscript.Interactive = True	
	On Error Resume Next
		
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\", "", "REG_SZ"		
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\", "", "REG_SZ"		
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\default path", "c:\temp", "REG_SZ"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\regday", "9718B059CEDB26D475CE27790CFE998916AD36B0B077DD09", "REG_SZ"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SMTPAttachment", strUserName & "_PartsCatFile.pdf", "REG_SZ"	
	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\", "", "REG_SZ"
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\default path", "c:\temp", "REG_SZ"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\disable url detection", 0, "REG_DWORD"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\file options", 97, "REG_DWORD"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\file options", 1121, "REG_DWORD"		
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\last save type", 1, "REG_DWORD"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\last tab", 0, "REG_DWORD"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\PDFAuthor", strUserName, "REG_SZ"
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\PDFDocumentScaling", 100, "REG_DWORD"
	'objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\PDFMailRecipients", "mbarrett@powerfarming.co.nz", "REG_SZ"	
	'objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\PDFSubject", "Power Farming Ltd. - Parts Cat File Export", "REG_SZ"	
	'objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\PDFTitle", "PartsCatFile", "REG_SZ"		
	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\persistent", 1, "REG_DWORD"				
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\persistent", 1, "REG_DWORD"				
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\file options", 1121, "REG_DWORD"			
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\file options", 1121, "REG_DWORD"			
	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\SaveDialog\security permissions", 4294967292, "REG_DWORD"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\PDFMailServerName", "smtp.powerfarming.co.nz", "REG_SZ"
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\PDFMailPort", 25, "REG_DWORD"
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\PDFMailFrom", "noreply@powerfarming.co.nz", "REG_SZ"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\PDFSubject", "Power Farming Ltd. - Parts Cat File Export", "REG_SZ"	
	objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\PDFTitle", "PartsCatFile", "REG_SZ"	

	PDFMailRecipients = ""
	PDFMailRecipients = Trim(objWshShell.RegRead("HKCU\Software\Dane Prairie Systems\Win2PDF\PDFMailRecipients"))
	If PDFMailRecipients = "" Then
		objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\PDFMailRecipients", "", "REG_SZ"				
	End If
		
	EventLog("UserTypeDetection:" & GetCustomerType())		
		
	Select Case GetCustomerType()
		Case "Employee"
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-ACC")
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-ASS")  
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-COP")  	
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-DES")  
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-ACC")  
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-LOG")  
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-LOG1")  				
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-PRT")  
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-REC")  
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-SHP")  
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-WAR")     
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-IT1")
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-LP2")				
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-COL")
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-MGT")
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-WAR")
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-MVL-PB1")
				objNet.AddWindowsPrinterConnection("\\PFNZ-SRV-029\PFW-CHC-WHS")		
				If PDFMailRecipients = "" Then
					objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\PDFMailRecipients", strUserName & "@powerfarming.co.nz", "REG_SZ"								
				End If
		Case "Retail"
			'Get user object
			Set User = GetObject("WinNT://" & "powerfarming.co.nz" & "/" & strUserName & ",user")
			'Loop through all groups user is a member of
			
				For Each Group in User.Groups					
					'Do work based on group membership
					Select Case Group.Name
						Case "stdgrpAshburton"
						 Call GROUPJOB_stdgrp_ASHBURTON_2012
						Case "stdgrp_MABERMOTORS"
						 Call GROUPJOB_stdgrp_MABERMOTORS_2012
						Case "stdgrp_AgriLife"
						 Call GROUPJOB_stdgrp_AgriLife_2012
						Case "stdgrp_Training"
						 Call GROUPJOB_stdgrp_Training
						Case "stdgrp_PowerTrac"
						 Call GROUPJOB_stdgrp_HAWKESBAY_2012
						Case "stdgrp_PFMANAWATU"
						 Call GROUPJOB_stdgrp_MANAWATU_2012
						Case "stdgrp_PFAGSOUTHLAND"
						 Call GROUPJOB_stdgrp_INVERCARGILL_2012
						Case "stdgrp_MABERTRACTORS"
						 Call GROUPJOB_stdgrp_TEAWAMUTU_2012
						Case "stdgrp_AgEarth"
						 Call GROUPJOB_stdgrp_NORTHLAND_2012
						Case "stdgrp_BROWNWOODS"
						 Call GROUPJOB_stdgrp_TIMARU_2012
						Case "stdgrp_CanterburyTractors"
						 Call GROUPJOB_stdgrp_CANTERBURY_2012
						Case "stdgrp_PREMIER"
						 Call GROUPJOB_stdgrp_TARANAKI_2012
						Case "stdgrpOTAGO"
						 Call GROUPJOB_stdgrp_OTAGO_2012
						Case "stdgrp_GISBORNE"
						 Call GROUPJOB_stdgrp_GISBORNE_2012
						Case "stdgrp_HAMILTON"
						 Call GROUPJOB_stdgrp_AGRILIFE_2012
					End Select			    
				Next		
				If PDFMailRecipients = "" Then
					objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\PDFMailRecipients", strUserName & "@powerfarming.co.nz", "REG_SZ"								
				End If				
		Case "Dealer"
			objNET.SetDefaultPrinter "Win2PDF"	
				If PDFMailRecipients = "" Then
					objWshShell.RegWrite "HKCU\Software\Dane Prairie Systems\Win2PDF\PDFMailRecipients", strUserName & "@powerfarming.co.nz", "REG_SZ"								
					retval = objWshShell.Run ("C:\WINDOWS\system32\mshta.exe" & Chr(34) & " C:\Program Files\PartsCatEmailAddress\PartsCatEmailAddress.hta" & Chr(34), 0, FALSE)					
				End If									
		Case Else
	End Select

	If err.number <> 0 then
		EventLog("PartsCat_SetupLocalResources|" & err.number & "|" & err.description)
	End If	
	
End Sub				

Sub EventLog(strLog)
	Err.Clear
	On Error Resume Next	
	objWshShell.LogEvent 8, "PartsCatAppLogin:" & strUserName & ":" & strLog  
  On Error Goto 0
  Err.Clear
End Sub

				