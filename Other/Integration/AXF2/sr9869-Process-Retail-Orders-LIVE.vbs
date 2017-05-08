
Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objNet = Wscript.CreateObject("Wscript.Network")
Wscript.Interactive = false

Public dtStartNow
Public dtLastNow

Call Main

Sub Main


    'Set Start Time
    dtStartNow = Now
    dtLastNow = Now
		
	Call StartCheck
	Call LoopAccount
	Call EndCheck

End Sub	
	
	
Sub LoopAccount
    Set cn = CreateObject("ADODB.Connection")
    cn.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-019\PFWAX" & ";Initial Catalog=" & "PFW_AX2009_LIVE" & ";Integrated Security=SSPI;"
    Set rs = CreateObject("ADODB.Recordset")
    rs.ActiveConnection = cn
	sql = "SELECT CustAccount FROM Datamart.[dbo].[sr9869 - GetCustomerAccounts]('pfw')"
    rs.Open sql,cn,1,1
    'rs.MoveFirst
	
	If rs.BOF and rs.EOF then
		Log("LoopAccount:No Accounts Found.")		
	Else
		Do Until rs.EOF
			Log("LoopAccount:Processing Customer " & rs("CustAccount").Value)
			LoopUpdateOrders(rs("CustAccount").Value)
			rs.MoveNext
		Loop	
	End If	
	
	Wscript.Echo "Done"
	

End Sub

Sub LoopUpdateOrders(custaccount)

	On Error Resume Next
	Log("	LoopUpdateOrders Starts")

	Const adParamInput = 1
	Const adParamOutput = 2
	Const adParamInputOutput = 3
	Const adParamReturnValue = 4
	Const adCmdStoredProc = &H0004

	Const adChar = 129
	Const adVarChar = 200
	Const adInteger = 3
	
	Set cmd = CreateObject("ADODB.Command")
	With cmd
		.ActiveConnection = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-019\PFWAX" & ";Initial Catalog=" & "DataMart" & ";Integrated Security=SSPI;CommandTimeout=1000"
		.CommandType = adCmdStoredProc
		.CommandText = "[dbo].[sr9869 - GetSnapOrdersToUpdate]"
		.CommandTimeout = 180
		.Parameters.Refresh	 
		.Parameters(1) = custaccount
		.Parameters(2) = GetCustomerParameter("pfw", custaccount, "SourceEndPointLocalStartDateTime")
	End With	
	Set rs = cmd.Execute
	
	If rs.BOF and rs.EOF then
		Log("	LoopUpdateOrders:No Orders for Customer " & custaccount)
	Else
		Do Until rs.EOF		
			Log("	LoopUpdateOrders:Processing Order " & rs(0).Value & " for Customer " & custaccount)
			Call SaveXmlFile(custaccount, rs(0).Value, GetXMLOrder("pfw", rs(0).Value, custaccount))
			rs.MoveNext
		Loop  
	End If
	
	If Err.Number <> 0 Then
		Log("	Error: " & Err.Number & " " & Err.Description)
	End If
	
	set cmd = nothing
	set rs = nothing
	
	Log("	LoopUpdateOrders Ends")

End Sub

Sub SaveXMLFile(custaccount, salesid, xml)
			
	If InStr(xml, "InvoiceID") <> 0 Then
		Log("		Order " & salesid & " appears to have one or more invoices.")
	End If
	If InStr(xml, "Header") = 0 Then
		Log("		Error found with Order " & salesid & ". It appears to be missing an XML header.")
	End If
	fn = GetCustomerParameter("pfw", custaccount, "PathInitialDrop") & "\" & "F2-" & Replace(custaccount,"/","-") & "-PurchaseOrder-" & salesid & "-AX-" & TimeStamp & ".xml"
    Set f = objFSO.CreateTextFile(fn, True)	
    f.Write xml
    f.Close
	Log("		Order " & salesid & " written to file: " & fn)
	
End Sub

Function GetXmlOrder(dataareaid, salesid, custaccount)

	Const adParamInput = 1
	Const adParamOutput = 2
	Const adParamInputOutput = 3
	Const adParamReturnValue = 4
	Const adCmdStoredProc = &H0004

	Const adChar = 129
	Const adVarChar = 200
	Const adInteger = 3
	
	Set cmd = CreateObject("ADODB.Command")
	With cmd
		.ActiveConnection = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-019\PFWAX" & ";Initial Catalog=" & "DataMart" & ";Integrated Security=SSPI;"
		.CommandType = adCmdStoredProc
		.CommandText = "[dbo].[sr9869 - GetXMLOrder]"
		.CommandTimeout = 600
		.Parameters.Refresh	 
		.Parameters(1) = dataareaid
		.Parameters(2) = salesid
		.Parameters(3) = custaccount
	End With
	
	Set rs = cmd.Execute	
    rs.MoveFirst
	GetXmlOrder = rs(0).Value

	set cmd = nothing
	set rs = nothing
	
End Function

Function GetCustomerParameter(dataareaid, custaccount, parameter)

	Const adParamInput = 1
	Const adParamOutput = 2
	Const adParamInputOutput = 3
	Const adParamReturnValue = 4
	Const adCmdStoredProc = &H0004

	Const adChar = 129
	Const adVarChar = 200
	Const adInteger = 3
	
	Set cmd = CreateObject("ADODB.Command")
	With cmd
		.ActiveConnection = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-019\PFWAX" & ";Initial Catalog=" & "DataMart" & ";Integrated Security=SSPI;"
		.CommandType = adCmdStoredProc
		.CommandText = "[dbo].[sr9869 - GetCustHeaderParameter]"
		.Parameters.Refresh	 
		.Parameters(1) = dataareaid
		.Parameters(2) = custaccount
		.Parameters(3) = parameter		
	End With
	
	Set rs = cmd.Execute	
    rs.MoveFirst
	GetCustomerParameter = rs(0).Value

	set cmd = nothing
	set rs = nothing
	
End Function

Function HomePath
	HomePath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
End Function

Function LZ(ByVal Number)
  If Number < 10 Then
    LZ = "0" & CStr(Number)
  Else
    LZ = CStr(Number)
  End If
End Function

Function TimeStamp
  Dim CurrTime
  CurrTime = Now()

  TimeStamp = CStr(Year(CurrTime)) & "" _
    & LZ(Month(CurrTime)) & "" _
    & LZ(Day(CurrTime)) & "" _
    & LZ(Hour(CurrTime)) & "" _
    & LZ(Minute(CurrTime)) & "" _
    & LZ(Second(CurrTime))
End Function

Sub StartCheck
	Err.Clear
	On Error Resume Next	

	Log("Script Starts")
	
	If objFSO.FolderExists("c:\Support") <> True Then	
		objFSO.CreateFolder "c:\Support"
	End If
	
	fl = "C:\Support\sr9869-Process-Retail-Orders-LIVE.chk"	
	If objFSO.FileExists(fl) = True Then
		Log("	Error: DirtyShutdownDetected")
		objFSO.DeleteFile fl, True
	End If
	
	Set fileLogon = objFSO.OpenTextFile(fl, 8, True)	
	fileLogon.Write "."
	fileLogon.WriteLine
  
	fileLogon.Close
	On Error Goto 0
	Err.Clear
End Sub

Sub EndCheck
	Err.Clear
	On Error Resume Next	
	Log("Script Ends")	
	fl = "C:\Support\sr9869-Process-Retail-Orders-LIVE.chk"	
	objFSO.DeleteFile fl, True

	On Error Goto 0
	Err.Clear
End Sub

Sub Log(strLog)
	Err.Clear
	On Error Resume Next	

	If objFSO.FolderExists("c:\Support") <> True Then	
		objFSO.CreateFolder "c:\Support"
	End If
	
	Set fileLogon = objFSO.OpenTextFile("C:\Support\sr9869-Process-Retail-Orders-LIVE-log.txt", 8, True)

	dtJustNow = Now
	sLastLog = DateDiff("s", dtLastNow, dtJustNow)
	
	fileLogon.Write dtJustNow & " (" & sLastLog & "secs) " & strLog
	fileLogon.WriteLine
	dtLastNow = Now

  fileLogon.Close

  On Error Goto 0
  Err.Clear
End Sub


