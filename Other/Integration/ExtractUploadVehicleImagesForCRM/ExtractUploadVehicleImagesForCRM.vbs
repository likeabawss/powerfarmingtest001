

Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objNet = Wscript.CreateObject("Wscript.Network")

'	1. Checks an HTTP location for:
'		a) a required xml file
'		b) a bunch of jpg files
'	2. Downloads ( per branch ) the XML file, and from that grabs the graphic with the assoc. listing information from the XML file.
'		Initially stores all of the branch JPG in a base location.
'	3. Subsequently, downloads the files again to TEMP, one-by-one and compares size in bytes and the file name
'		with the existing files.
'		a) if there is a file name match, but the stored bytes differ, puts a copy in the /updated folder, and overwrites the original. ( ready to be compared next time. )
'		b) if there is a file name match, and the stored bytes are the same, deletes the newly downloaded file and does nothing else.
'		c) if there is no file match because the base file is missing, stores a copy in the /updated folder, and creates a base. ( ready to be compared next time. )
'				- this is probably a new listing, or an existing listing that did not previously have an initial _1 image.
'
'	Notes:
'	1. The TEMP folder is cleared at each run.
'	2. The images in the Branch Update folder are retained and are refreshed from TEMP at each run.
'

'Init Pubs
stagingFolder = "C:\Support\FastTrack"
branchBaseImageFolder = ""
branchUpdatedImageFolder = ""
branchUpdatedTempFolder = ""
dtStartNow = Now
dtLastNow = Now

Call Main

Sub Main

	Log("Script Starts: " & objFSO.GetFile(Wscript.ScriptFullName))

	Dim Branchlist(1)
	Branchlist(0) = "Waikato"
	Branchlist(1) = "ashburton"
			
	On Error Resume Next
	objFSO.DeleteFile(stagingFolder & "\*.xml" )
	On Error Goto 0
		
	'Download all images to Branch\Temp
	For Each Branch In Branchlist
		Call InitialiseBranchFolderStructure(branch)
		pathToXML = "http://usedproductws-prod.powerfarming.co.nz/data/" & Branch & ".xml"
		Log("Processing Branch: " & Branch)
		Log("	pathToXML: " & pathToXML)
		
		Call HTTPDownload( pathToXML, stagingFolder )
		If objFSO.FileExists(stagingFolder & "\" & Branch & ".xml") Then
			Call DownloadAllImages(branch, stagingFolder & "\" & Branch & ".xml")			
		Else
			Log("	XML download failure for: " & pathToXML)
		End If
		
		Call ProcessImages()
	Next

	Log("Script Ends: " & objFSO.GetFile(Wscript.ScriptFullName))
	
End Sub

'Create Required Folders
Sub InitialiseBranchFolderStructure(branch)

	branchBaseImageFolder = stagingFolder & "\" & branch
	If objFso.FolderExists(branchBaseImageFolder) = False Then
		objFSO.CreateFolder(branchBaseImageFolder)		
	End If
	
	branchUpdatedImageFolder = stagingFolder & "\" & branch & "\Updated"
	If objFso.FolderExists(branchUpdatedImageFolder) = False Then		
		objFSO.CreateFolder(branchUpdatedImageFolder)
	End If

	branchUpdatedTempFolder = stagingFolder & "\" & branch & "\Temp"
	If objFso.FolderExists(branchUpdatedTempFolder) = False Then		
		objFSO.CreateFolder(branchUpdatedTempFolder)
	Else
		objFSO.DeleteFile(branchUpdatedTempFolder & "\*.*" )
	End If
	
End Sub

'Iterate listings in XML, acquire all images to TEMP.
Sub DownloadAllImages(branch, xmlFl)

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.Async = "false"
	xmlDoc.Load(xmlFl)	
	Set objNodeList = xmlDoc.getElementsByTagName("listing")
	
	If objNodeList.length > 0 then
		For Each listing in xmlDoc.SelectNodes("//listing")
			listingId = listing.getAttribute("id")	
			stocknumber = listing.SelectSingleNode("stock_number").Text					
			imgUrl = "http://usedproductws-prod.powerfarming.co.nz/data/" & branch & "_" & listingId & "_1.jpg"
			Call HTTPDownload( imgUrl, branchUpdatedTempFolder )
			If objFSO.FileExists(branchUpdatedTempFolder & "\" & branch & "_" & listingId & "_1.jpg") Then
				newfilename = branchUpdatedTempFolder & "\" & stocknumber & ".jpg" 
				objFSO.MoveFile branchUpdatedTempFolder & "\" & branch & "_" & listingId & "_1.jpg", newfilename
				Log("		" & branch & "_" & listingId & "_1.jpg" & " downloaded, renamed to " & newfilename)
			Else
				Log("		" & branch & "_" & listingId & "_1.jpg" & " FAILED to download.")
			End If				
		Next
	End If
		
End Sub

Sub ProcessImages()

	Log("	Processing Images in: " & branchUpdatedTempFolder)
	wscript.echo branchUpdatedTempFolder
	Set fls = objFSO.GetFolder(branchUpdatedTempFolder).Files
	For Each fl in fls
		If objFSO.FileExists(branchBaseImageFolder & "\" & fl.Name) Then
			If fl.Size = objFSO.GetFile(branchBaseImageFolder & "\" & fl.Name).Size Then				
				Log("		(Check) No difference detected. Deleted " & fl)
				fl.Delete
			Else				
				Log("		(Check) Difference found. Replacing base with new and queuing " & fl.Name & " for CRM update.")				
				objFSO.CopyFile fl, branchUpdatedImageFolder & "\" & fl.Name
				objFSO.CopyFile fl, branchBaseImageFolder & "\" & fl.Name
				fl.Delete
			End If
		Else
			flnm = fl.Name
			objFSO.CopyFile fl, branchUpdatedImageFolder & "\" & fl.Name
			objFSO.MoveFile fl, branchBaseImageFolder & "\" & flnm
			If objFSO.FileExists(branchBaseImageFolder & "\" & flnm) Then
				Log("		(New) Moved " & flnm & " to " & branchBaseImageFolder & " and queued for update.")
			Else
				Log("		(New) Move FAIL " & flnm & " to " & branchBaseImageFolder)
			End If
		End If
		'wscript.echo fl
	Next
		
End Sub

Sub HTTPDownload( myURL, myPath )

	On Error Resume Next

    ' Standard housekeeping
    Dim i, objFile, objFSO, objHTTP, strFile, strMsg
    Const ForReading = 1, ForWriting = 2, ForAppending = 8

    ' Create a File System Object
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ' Check if the specified target file or folder exists,
    ' and build the fully qualified path of the target file
    If objFSO.FolderExists( myPath ) Then
        strFile = objFSO.BuildPath( myPath, Mid( myURL, InStrRev( myURL, "/" ) + 1 ) )
    ElseIf objFSO.FolderExists( Left( myPath, InStrRev( myPath, "\" ) - 1 ) ) Then
        strFile = myPath
    Else
        WScript.Echo "ERROR: Target folder not found."
        Exit Sub
    End If

    ' Create or open the target file
    Set objFile = objFSO.OpenTextFile( strFile, ForWriting, True )

    ' Create an HTTP object
    Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )

    ' Download the specified URL
    objHTTP.Open "GET", myURL, False
    objHTTP.Send

    ' Write the downloaded byte stream to the target file
    For i = 1 To LenB( objHTTP.ResponseBody )
        objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
    Next

    ' Close the target file
    objFile.Close( )
End Sub

Sub Log(strLog)

	Err.Clear
	On Error Resume Next	
	
	Set fileLogon = objFSO.OpenTextFile(objFSO.GetParentFolderName(objFSO.GetFile(Wscript.ScriptFullName)) & "\" & Replace(Wscript.ScriptName, ".vbs", ".log"), 8, True)

	dtJustNow = Now
	sLastLog = DateDiff("s", dtLastNow, dtJustNow)
	
	fileLogon.Write dtJustNow & " (" & sLastLog & "secs) " & strLog
	fileLogon.WriteLine
	dtLastNow = Now

	fileLogon.Close

	On Error Goto 0
	Err.Clear
	
End Sub
