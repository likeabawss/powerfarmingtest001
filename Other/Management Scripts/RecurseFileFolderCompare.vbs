' Objective: To delete old files from a given folder and all subfolders below
' Created by: MAK
' Created Date: June 21, 2005
' Usage: cscript deloldfiles.vbs c:\dba\log 3
'      : It deletes files older than 3 days


'
'	Logged Deletion Script
'	**********************
'		by Mike Barrett
'	
'	1. To recursively delete and log those deletions of files that are older than a chosen age.
'	2. To compare a source folder with a destination, and delete ( while logging ) files that are confirmed to exist
'		in the destinatino.


Wscript.Interactive = FALSE

Set objArgs = WScript.Arguments
FolderName =objArgs(0)
Days=objArgs(1)
 
set fso = createobject("scripting.filesystemobject")
set folders = fso.getfolder(FolderName)
datetoday = now()
newdate = dateadd("d", Days*-1, datetoday)
wscript.echo "Today:" & now()
wscript.echo "Started deleting files older than :" & newdate 
wscript.echo "________________________________________________"
wscript.echo ""
recurse folders 
wscript.echo ""
wscript.echo "Completed deleting files older than :" & newdate 
wscript.echo "________________________________________________"
 
sub recurse( byref folders)
  set subfolders = folders.subfolders
  set files = folders.files
  wscript.echo ""
  wscript.echo "Deleting Files under the Folder:" & folders.path
  wscript.echo "__________________________________________________________________________"
  for each file in files
    if file.datelastmodified < newdate then
      wscript.echo "Deleting " & folders.path & "\" & file.name & " last modified: " & file.datelastmodified
      on error resume next
    file.delete
    end if
 
  next  
 
  for each folder in subfolders
    recurse folder
  next  
 
  set subfolders = nothing
  set files = nothing
 
end sub