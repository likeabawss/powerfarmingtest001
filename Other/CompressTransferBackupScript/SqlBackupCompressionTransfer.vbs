
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'
'																	File Backup Compression Transfer
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'
'	ChangeLog:
'	**********
'	
'	ToDo:
'	*****
'	14012013 - MB - Log events to local windows event logger.
'	14012013 - MB - Log failed compression / transfer events to windows event logger.
'
'

	'Load Intrinsic Objects
	Set objWshShell = Wscript.CreateObject("Wscript.Shell")
	Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
	Set objNet = Wscript.CreateObject("Wscript.Network")

	Dim sLogFile
	Dim logSqlServer
	Dim logSqlDatabase
	Dim bkpSystem
	
	bkpSystem = WScript.Arguments(0)
	compressPath = WScript.Arguments(1)
	transferPath = WScript.Arguments(2)
	cdlFileExtensions = WScript.Arguments(3)
	logSqlServer = WScript.Arguments(4)
	logSqlDatabase = WScript.Arguments(5)
	logFile = WScript.Arguments(6)
	deleteAfterTransfer = WScript.Arguments(7)	
		
	If deleteAfterTransfer <> 0 Then
		deleteAfterTransfer = True
	End If
		
    'Set Start Time
    dtStartNow = Now
    dtLastNow = Now
		
	Call DoCompressionTransferRecurse(compressPath, transferPath, cdlFileExtensions, logSqlServer, logSqlDatabase, logFile, deleteAfterTransfer)
	
	Log("SqlBackupCompressionTransfer: Ends")			
		
	Sub DoCompressionTransferRecurse(compressPath, transferPath, cdlFileExtensions, logSqlServer, logSqlDatabase, logFile, deleteAfterTransfer)
		'Replicate Folder Structure	
		objWshShell.Run "cmd /c robocopy.exe " & Chr(34) & compressPath & Chr(34) & " " & Chr(34) & transferPath & Chr(34) & " /E /Z /XF *", 2, True
		
		
		sLogFile = logFile
		Log("SqlBackupCompressionTransfer: Starts")
		Log("Processing of folder: " & compressPath & " and subsequent transfer to " & transferPath & " for file extensions " & cdlFileExtensions)		
		Set compressFolders = objFSO.getfolder(compressPath)
		Call Recurse(compressFolders, compressPath, transferPath, cdlFileExtensions, deleteAfterTransfer)
		Log("Completed processing of folder: " & compressPath & " and subsequent transfer to " & transferPath & " for file extensions " & cdlFileExtensions)	
	End Sub

	Function IsFileExtensionOk(fileName, commadelExtensionList)	
		commadelExtensionList = Replace(commadelExtensionList, ".", "")
		arExtensionList = Split(commadelExtensionList, ",")
		For i = LBound(arExtensionList) to UBound(arExtensionList)
			If UCase(GetFileExtension(fileName)) = arExtensionList(i) Then
				IsFileExtensionOk = True
				Exit Function
			End If
		Next		
		IsFileExtensionOk = False
	End Function

	Function fileClosed(fileObject)	
		On Error Resume Next
		Do While Left(fileObject.Name, 1) = "_"
			fileObject.Name = Mid(fileObject.Name, 2)
		Loop
		post = fileObject.ParentFolder & "\_" & fileObject.Name
		fileObject.Name = "_" & fileObject.Name
		If objFSO.FileExists(post) Then
			fileClosed = True
		Else
			fileClosed = False
		End If	

		Do While Left(fileObject.Name, 1) = "_"
			fileObject.Name = Mid(fileObject.Name, 2)
		Loop		
		
		On Error Goto 0	
	End Function
	
	Sub Recurse(byref folders, compressPath, transferPath, cdlFileExtensions, deleteAfterTransfer)	  
		On Error Resume Next
	  Set subfolders = folders.subfolders
	  Set files = folders.files
	  For Each file In files
		If IsFileExtensionOk(file.Name, cdlFileExtensions) Then
			If fileClosed(file) Then
			
				Log("Processing file " & file.Name)
				localCompressedFilePath = ""
				remoteTransferredFilePath = ""
										
				localCompressedFilePath = Replace(file.Path, "." & GetFileExtension(file.Path), ".7z")		
				remoteTransferredFilePath = Replace(localCompressedFilePath, compressPath, transferPath, 1, -1, 1)
				remoteTransferredFilePathParentFolder = Replace(file.ParentFolder, compressPath, transferPath, 1, -1, 1)
				
				'Check for local compressed file.
				If objFSO.FileExists(localCompressedFilePath) = False Then
					'Check for transferred compressed copy.
					If objFSO.FileExists(remoteTransferredFilePath) = False	Then		
					
						Call dbLogNewEntry(logSqlServer, logSqlDatabase, bkpSystem, file.ParentFolder, remoteTransferredFilePathParentFolder, file.Name, Replace(file.Name, "." & GetFileExtension(file.Name), ".7z"))				
						Call dbLogChangeStatus(logSqlServer, logSqlDatabase, file.Name, "CompressionStart")
						
						'Compress									
						If Zip(file.Path, localCompressedFilePath) = 1 Then
							Call dbLogChangeStatus(logSqlServer, logSqlDatabase, file.Name, "CompressionEnd")											
							'Transfer
							Call dbLogChangeStatus(logSqlServer, logSqlDatabase, file.Name, "TransferStart")											
							objFSO.MoveFile localCompressedFilePath, remoteTransferredFilePath
							If objFSO.FileExists(remoteTransferredFilePath)	Then		
									Call dbLogChangeStatus(logSqlServer, logSqlDatabase, file.Name, "TransferEnd")											
									Log("File " & file.Name & " was compressed and transferred to " & remoteTransferredFilePath & " sucessfully.")
									Call dbLogChangeStatus(logSqlServer, logSqlDatabase, file.Name, "Complete")
									If deleteAfterTransfer Then
										file.Delete True
									End If
							Else
								Log("Transfer Error - File " & file.Name & " failed To process.")
							End If
						Else
							Log("Zip Error - File " & file.Name & " failed To process.")				
						End If
					Else
						Log("File " & remoteTransferredFilePath & " already exists remotely. Skipped.")
					End If
				Else
					Log("File " & localCompressedFilePath & " already exists locally. Skipped.")
				End If
			Else
				Log("File " & localCompressedFilePath & " is currently locked. Skipped.")
			End If
		End If
	  Next  
	 
	  For Each folder In subfolders
		recurse folder, compressPath, transferPath, cdlFileExtensions, deleteAfterTransfer
	  Next  
	 
	  Set subfolders = Nothing
	  Set files = Nothing
	 
	End Sub

	Function GetFileExtension(fileName)
		lastperiodPos = InStr(StrReverse(fileName), ".")
		If lastperiodPos = 0 Then
			GetFileExtension = ""
		Else
			GetFileExtension = Right(fileName,(lastperiodPos-1))	
		End If
	End Function

	Function Zip(sFile,sArchiveName)
	  'This script is provided under the Creative Commons license located
	  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
	  'be used for commercial purposes with out the expressed written consent
	  'of NateRice.com   

	  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	  Set oShell = WScript.CreateObject("Wscript.Shell")

	  '--------Find Working Directory--------
	  aScriptFilename = Split(Wscript.ScriptFullName, "\")
	  sScriptFilename = aScriptFileName(Ubound(aScriptFilename))
	  sWorkingDirectory = Replace(Wscript.ScriptFullName, sScriptFilename, "")
	  '--------------------------------------

	  '-------Ensure we can find 7z.exe------
	  If objFSO.FileExists(sWorkingDirectory & "\" & "7z.exe") Then
		s7zLocation = ""
	  ElseIf objFSO.FileExists("C:\Program Files\7-Zip\7z.exe") Then
		s7zLocation = "C:\Program Files\7-Zip\"
	  Else
		Zip = "Error: Couldn't find 7z.exe"
		Exit Function
	  End If
	  '--------------------------------------
		
	  Log("""" & s7zLocation & "7z.exe"" a -y -mx1 """ & sArchiveName & """ " & """" & sFile & """")
	  
	  oShell.Run """" & s7zLocation & "7z.exe"" a -y -mx1 """ & sArchiveName & """ " & """" & sFile & """", 0, True   		

	  If objFSO.FileExists(sArchiveName) Then
		Zip = 1
		Log("Compression of " & sArchiveName & " was successful.")
	  Else
		Zip = "Error: Archive Creation Failed."
		Log("Compression of " & sArchiveName & " was FAILED.")
	  End If
	  
	End Function

	Function UnZip(sArchiveName,sLocation)
	  'This script is provided under the Creative Commons license located
	  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
	  'be used for commercial purposes with out the expressed written consent
	  'of NateRice.com

	  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	  Set oShell = WScript.CreateObject("Wscript.Shell")

	  '--------Find Working Directory--------
	  aScriptFilename = Split(Wscript.ScriptFullName, "\")
	  sScriptFilename = aScriptFileName(Ubound(aScriptFilename))
	  sWorkingDirectory = Replace(Wscript.ScriptFullName, sScriptFilename, "")
	  '--------------------------------------

	  '-------Ensure we can find 7z.exe------
	  If objFSO.FileExists(sWorkingDirectory & "\" & "7z.exe") Then
		s7zLocation = ""
	  ElseIf objFSO.FileExists("C:\Program Files\7-Zip\7z.exe") Then
		s7zLocation = "C:\Program Files\7-Zip\"
	  Else
		UnZip = "Error: Couldn't find 7z.exe"
		Exit Function
	  End If
	  '--------------------------------------

	  '-Ensure we can find archive to uncompress-
	  If Not objFSO.FileExists(sArchiveName) Then
		UnZip = "Error: File Not Found."
		Exit Function
	  End If
	  '--------------------------------------

	  oShell.Run """" & s7zLocation & "7z.exe"" e -y -o""" & sLocation & """ """ & _
	  sArchiveName & """", 0, True
	  UnZip = 1
	End Function

	Sub Log(strLog)
		Err.Clear
		On Error Resume Next	
		
		'Wscript.Echo strLog
		
		'Open / Create Text File
		Set fileLogon = objFSO.OpenTextFile(sLogFile, 8, True)

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
	
	Sub dbLogChangeStatus(logSqlServer, logSqlDatabase, bkpSourceFile, bkpStatus)
	
		On Error Resume Next
		Set cn = CreateObject("ADODB.Connection")
		cn.open = "Provider=SQLOLEDB.1; Data Source=" & logSqlServer & ";Initial Catalog=" & logSqlDatabase & ";Integrated Security=SSPI;"

		Set rs = CreateObject("ADODB.Recordset")
		rs.ActiveConnection = cn
		bkpSourceFile = Mid(bkpSourceFile, 1, (InStr(bkpSourceFile, ".")-1))
		sql = "Execute [dbo].[ChangeStatusSqlBackupCompressTransfer] '" & bkpStatus & "', '" & bkpSourceFile & "'"					
		rs.Open sql,cn,1,1
		On Error Goto 0
		err.clear		
	
	End Sub	
	
	
	Sub dbLogNewEntry(logSqlServer, logSqlDatabase, bkpSystem, bkpSourceFilePath, bkpDestinationFilePath, bkpSourceFile, bkpDestinationFile)
				
		'On Error Resume Next
		Set cn = CreateObject("ADODB.Connection")
		cn.open = "Provider=SQLOLEDB.1; Data Source=" & logSqlServer & ";Initial Catalog=" & logSqlDatabase & ";Integrated Security=SSPI;"			

		Set rs = CreateObject("ADODB.Recordset")
		rs.ActiveConnection = cn
		sql = "Execute [dbo].[NewSqlBackupCompressTransferMaster] '" & bkpSystem & "', '" & _
				bkpSourceFilePath & "', '" & bkpDestinationFilePath & "', '" & bkpSourceFile & "', '" & bkpDestinationFile & "'"
		rs.Open sql,cn,1,1
		On Error Goto 0
		err.clear		
	
	End Sub