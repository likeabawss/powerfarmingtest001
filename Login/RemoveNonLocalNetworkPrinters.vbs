
'Load Intrinsic Objects
Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objNet = Wscript.CreateObject("Wscript.Network")

Dim strLocation
Call GetLocation
Call RemoveNonLocalNetworkPrinters

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
	   If Instr(strIPAddress, "192.168.201.") <> 0 OR _
              Instr(strIPAddress, "192.168.203.") <> 0 OR _
              Instr(strIPAddress, "161.71.70.") <> 0 OR _
              Instr(strIPAddress, "192.168.206.") <> 0 OR _
              Instr(strIPAddress, "192.168.3.") <> 0 OR _
              Instr(strIPAddress, "192.168.0.") <> 0 Then 
		   GetIP = strIPAddress
	   End If
	End If
     Next
  Next

  'Disable Error Handling
  On Error Goto 0

End Function

Sub GetLocation

  'Clear Err Object
  Err.Clear

  'Enable Error Handling
  On Error Resume Next

  'GetIP
  strIP = GetIP

  'Set Location
  If Instr(strIP, "161.71.70.") <> 0 OR Instr(strIP, "192.168.99.") <> 0 Then
     'Set Location
     strLocation = "PFNZ"
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
  If Instr(strIP, "192.168.3.") <> 0 Then
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

  'Disable Error Handling
  On Error Goto 0

End Sub

Sub RemoveNonLocalNetworkPrinters

  'Enable Error Handling
  On Error Resume Next

  Select Case strLocation
    Case "PFNZ"
		exclusionText(0) = "PFG"
		exclusionText(1) = "HOW"
		'ipmask = "PFG"
    Case "PFGAU_MAIN"
		exclusionText(0) = "PFW"
		exclusionText(1) = "HOW"	
		'ipmask = "192.168.203."
    Case "PFGAU_SERVICE"
	
	 ipmask = "192.168.201."
    Case "PFNZ_MABERS"
	 ipmask = "192.168.3.;161.71.70."
	 Wscript.Quit
    Case "PFGAU_BRISBANE"
	 ipmask = "192.168.206."
    Case "HOWARD_SYDNEY"
	 ipmask = "192.168.0."
  End Select
  
  
  
  Set oPrinters = objNet.EnumPrinterConnections

  For i = 0 to oPrinters.Count - 1 Step 2
    If InStr(oPrinters.Item(i+1), "\\") <> 0 Then
      If InStr(oPrinters.Item(i), ipmask) = 0 Then
		 'WScript.Echo oPrinters.Item(i) & " " & oPrinters.Item(i+1)
		 objNet.RemovePrinterConnection oPrinters.Item(i+1), TRUE, TRUE
      End if
    End If
  Next

  'Disable Error Handling
  On Error Goto 0

End Sub
