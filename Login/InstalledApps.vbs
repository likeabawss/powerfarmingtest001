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
		  
	 If InStr(UCase(objItem.ProductName), UCase("powershell")) <> 0 Then
		If Left(objItem.Version, 3) = "1.0" Then		
	      MSISoftwareInstalled = TRUE
		  'wscript.echo objItem.ProductName
		  Wscript.Echo "Name: " & UCase(objItem.ProductName) & " Version:" & objItem.Version & " ID Number:" & objItem.IdentifyingNumber
	      Exit For
		End If
	 End If
     Next
  Next