

'**********************************************************************************************
'
'				   TableDiffBatchScipt
'
'**********************************************************************************************

' Author: Michael Barrett
' Date: 03/07/2014
'
'ChangeLog:
'**********
'
'
'
'

    'Load Intrinsic Objects
    Set objWshShell = Wscript.CreateObject("Wscript.Shell")
    Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objNet = Wscript.CreateObject("Wscript.Network")

	Dim ExportFolder
	Dim TableDiffExecutablePath
	Dim logSQLServer
	Dim logSQLDatabase
	Dim logSQLTable

	On Error Resume Next
    dtStartNow = Now
    dtLastNow = Now
	argXML = WScript.Arguments(0)
	parentFolder = objFSO.GetFile(WScript.ScriptFullName).ParentFolder
	
	Call Main

	Sub Main
	
		On Error Goto 0
			
		If argXML = "" Then
			WScript.Echo "Please specify required XML file location."
			WScript.Quit
		End If

		If objFSO.FileExists(argXML) = False Then
			If objFSO.FileExists(parentFolder & "\" & argXML) = True Then	
				argXML = parentFolder & "\" & argXML
			Else
				WScript.Echo "Unable to find required XML file. Please check file location."
				WScript.Quit
			End If	
		Else
			WScript.Echo "Unable to find required XML file. Please check file location."
			WScript.Quit		
		End If

		'Load XML
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")		
		xmlDoc.Async = "False"
		xmlDoc.Load(argXML)	
				
		'Load Conf
		Set colNodesExportFolder = xmlDoc.selectNodes("//TableDiffCheck/Conf")				
		For Each objNode in colNodesExportFolder
			ExportFolder = objNode.Attributes.getNamedItem("ExportFolder").Text
			TableDiffExecutablePath = objNode.Attributes.getNamedItem("TableDiffExecutablePath").Text
			logSQLServer = objNode.Attributes.getNamedItem("logSQLServer").Text
			logSQLDatabase = objNode.Attributes.getNamedItem("logSQLDatabase").Text
			logSQLTable = objNode.Attributes.getNamedItem("logSQLTable").Text
		Next
		
		'Validate Conf
		If objFSO.FolderExists(ExportFolder) = False Then
			errmsg = errmsg & "Export folder does not exist." & Chr(13) & Chr(10)
		End If			
		If objFSO.FileExists(TableDiffExecutablePath) = False Then
			errmsg = errmsg & "Could not find TableDiff executable." & Chr(13) & Chr(10)
		End If
		If errmsg <> "" Then
			WScript.Echo errmsg
			WScript.Quit
		End If
			
		WScript.Echo "ExportFolder: " & ExportFolder
		WScript.Echo "TableDiffExecutablePath: " & TableDiffExecutablePath
		WScript.Echo "logSQLServer: " & logSQLServer
		WScript.Echo "logSQLDatabase: " & logSQLDatabase
		WScript.Echo "logSQLTable: " & logSQLTable						
							
		'InjectJob
		
	
				
		'Load Commands				
		Set colNodesCheckTable = xmlDoc.selectNodes("//TableDiffCheck/Checks/CheckTable[@Enabled='Y']")		
		For Each objNode in colNodesCheckTable
		
			If objFSO.FileExists(exportFolder & "\" & objNode.Attributes.getNamedItem("SourceTable").Text & ".sql") Then
				objFSO.DeleteFile(exportFolder & "\" & objNode.Attributes.getNamedItem("SourceTable").Text & ".sql")
			End If	
		
			tdArgs = " -sourceserver " & objNode.Attributes.getNamedItem("SourceServer").Text & _
					 " -sourcedatabase " & objNode.Attributes.getNamedItem("SourceDatabase").Text & _
					 " -sourceschema " & objNode.Attributes.getNamedItem("SourceSchema").Text & _
					 " -sourcetable " & objNode.Attributes.getNamedItem("SourceTable").Text & _
					 " -destinationserver " & objNode.Attributes.getNamedItem("DestinationServer").Text & _ 
					 " -destinationdatabase " & objNode.Attributes.getNamedItem("DestinationDatabase").Text & _ 
					 " -destinationschema " & objNode.Attributes.getNamedItem("DestinationSchema").Text & _  
					 " -destinationtable " & objNode.Attributes.getNamedItem("DestinationTable").Text & _  
					 " -f " & Chr(34) & exportFolder & "\" & objNode.Attributes.getNamedItem("SourceTable").Text & ".sql" & Chr(34)
					 
		  	Call TableDiffRun(Chr(34) & TableDiffExecutablePath & Chr(34) & tdArgs)		  	
		  	If objNode.Attributes.getNamedItem("TextFilterInCSV").Text <> "" Then		  		
		  		UnMatchedRecs = 0
		  		UnMatchedRecs = FilterErroneousCount(exportFolder & "\" & objNode.Attributes.getNamedItem("SourceTable").Text & ".sql", objNode.Attributes.getNamedItem("TextFilterInCSV").Text)
		  	End If
		  	Call DBInjectResults(objNode.Attributes.getNamedItem("SourceTable").Text, objNode.Attributes.getNamedItem("SourceServer").Text, objNode.Attributes.getNamedItem("SourceDatabase").Text, objNode.Attributes.getNamedItem("SourceSchema").Text, objNode.Attributes.getNamedItem("DestinationTable").Text, objNode.Attributes.getNamedItem("DestinationServer").Text, objNode.Attributes.getNamedItem("DestinationDatabase").Text, objNode.Attributes.getNamedItem("DestinationSchema").Text, UnMatchedRecs)
		  	  	
		Next

	End Sub
	
	Function NewTableDiffJob
	
		sql = "INSERT INTO [dbo].[TableDiffCheckHeader] " & _
           	  "([XMLFile],[StartDateTime],[EndDateTime]) " & _
     			" VALUES('" & objFSO.GetFile(argXML).Name & "',getdate(),NULL,IsnUll((SELECT Max([Run])+1 FROM [dbo].[TableDiffCheckHeader] where XMLFile = '" & objFSO.GetFile(argXML).Name & "' group by XMLFile),0))"
     			
    	Set cn = CreateObject("ADODB.Connection")
	    cn.open = "Provider=SQLOLEDB.1; Data Source=" & logSQLServer & ";Initial Catalog=" & logSQLDatabase & ";Integrated Security=SSPI;"
		cn.Execute(sql)
		
    Set cn = CreateObject("ADODB.Connection")
    cn.open = "Provider=SQLOLEDB.1; Data Source=" & "PFNZ-SRV-028" & ";Initial Catalog=" & "iLient" & ";Integrated Security=SSPI;"

    Set rs = CreateObject("ADODB.Recordset")
    rs.ActiveConnection = cn
	sql = "select c.computer_name from dbo.computer c where c.computer_type = 'Server' and c.parent_group like '%POWERFARMING%'"
	
    rs.Open sql,cn,1,1
    rs.MoveFirst
	
	Do Until rs.EOF
		Wscript.Echo rs(0)
		rs.MoveNext
	Loop			
	
	End Fuction
	
	Sub TableDiffRun(cmd)
	
		Log("CommandLine:" & cmd)
		retval = objWshShell.Run (cmd, 0, True)
		WScript.Echo "cmd: " & cmd 		
	
	End Sub
	
	Function FilterErroneousCount(fl, includeLinesOnlyTextcsv)

		On Error Goto 0
		
		'WScript.Echo objFSO.GetFile(fl).ParentFolder
		'Wscript.Echo objFSO.GetBaseName(fl)
		'WScript.Echo objFSO.FileExists(objFSO.GetFile(fl).ParentFolder & "\" & objFSO.GetBaseName(fl) & "_Filtered.sql")
		
		If objFSO.FileExists(fl) Then
		
			If objFSO.FileExists(objFSO.GetFile(fl).ParentFolder & "\" & objFSO.GetBaseName(fl) & "_Filtered.sql") Then
				objFSO.DeleteFile(objFSO.GetFile(fl).ParentFolder & "\" & objFSO.GetBaseName(fl) & "_Filtered.sql")					
			End If
				
			 recs = 0
			 arIncl = Split(includeLinesOnlyTextcsv,",")	           
		     Set rd = objFSO.OpenTextFile(fl, 1)
		     Set wr = objFSO.OpenTextFile(objFSO.GetFile(fl).ParentFolder & "\" & objFSO.GetBaseName(fl) & "_Filtered.sql", 8, TRUE)
		     Do While rd.AtEndOfStream <> True			
				cLine = rd.ReadLine
				For Each mbr In arIncl
					If InStr(cLine, mbr) Then
						wr.WriteLine cLine
						recs = recs + 1
					End If				
				Next
		     Loop	     	     	     		 
				     
		     rd.Close
		     wr.Close
		     	     
		     FilterErroneousCount = recs
	     
	     Else
	    	FilterErroneousCount = -1
	    	WScript.Echo "No difference file found. Normally, there should be one, even if empty. File:" & kl 
	     End If 
	
		
	
	End Function
	
	Sub DBInjectResults(SourceTable, SourceServer, SourceDatabase, SourceSchema, DestinationTable, DestinationServer, DestinationDatabase, DestinationSchema, UnMatchedRecordCount)
	
		sql = "INSERT INTO [dbo].[TableDiffReplicationCheckResults] " & _
           		"([DateTime],[SourceTable],[SourceServer],[SourceDatabase],[SourceSchema]" & _
           		",[DestinationTable],[DestinationServer],[DestinationDatabase],[DestinationSchema]" & _
           		",[UnMatchedRecordCount]) " & _
			    "VALUES (GetDate(), '" & SourceTable & "','" & SourceServer & "','" & SourceDatabase & "','" & _
			    		 SourceSchema & "','" & DestinationTable & "','" & DestinationServer & "','" & _
			    		 DestinationDatabase & "','" & DestinationSchema & "'," & UnMatchedRecordCount & ")"
			    		 
		WScript.Echo sql			    		 
	
	    Set cn = CreateObject("ADODB.Connection")
	    cn.open = "Provider=SQLOLEDB.1; Data Source=" & logSQLServer & ";Initial Catalog=" & logSQLDatabase & ";Integrated Security=SSPI;"
		cn.Execute(sql)
		
	End Sub
	
	Sub Log(strLog)
	
		Err.Clear
		On Error Resume Next	
			
		'Open / Create Text File
		Set fileLogon = objFSO.OpenTextFile(parentFolder & "\" & WScript.ScriptName & "_log.txt", 8, True)
	
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
	
	
	