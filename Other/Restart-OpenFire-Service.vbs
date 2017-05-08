'============================================================
' RESTARTSERVICE - User Script
'------------------------------------------------------------
' This script is designed to permit end users to restart a stopped service.
'============================================================
' CREATED BY: rormeister - Spiceworks Member at large
'
' NOTES:
'   Call script from .CMD (via a shortcut on the users desktop)
'   User must have permission to the folder/file share on server
'   Modify Constants "insideQuotes" to suit your need
'   You will need the correct name of the service
'   Save .CMD and .VBS in a common folder
'   Create a .CMD file and enter the call 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'  echo off
'  cls
'  echo ...restarting (your process) Service
'  echo ...pausing 10 seconds
'  echo ...this window will close automatically
'  echo ...do not close manually
'  echo ...call Support if you have problems
'  echo   
'  cscript /nologo  \\SHARESERVER\SHAREFOLDER\RestartServices.vbs   
'  REM (note no spaces in path/file)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
' CAVEAT:
'   Author has not tested this specific code but all functionality has been left in tact.
'   Software is provided AS IS and the Author assumes no responsibility to it's suitability of purpose.
'   User is assumed to have working knowledge of VBScript, .CMD, and other Windows manipulation tools.
'
'=============================================================
' CHANGES:
'	07.27.2011 - Script Created, original script
'=============================================================

Option Explicit
ON Error GoTo 0

'// Constants
Const START_SVC = "Openfire"

'// Members
Dim objWMI

'//////////////////////////////////////////////////////////////////////////////
'// ENTRY POINT

' Sleep for 10 Seconds before we start
wScript.Sleep 10000

' Create our WMI object to retrieve the services
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 

' Restart Service - bisTrack eConnect Agent
Call RestartServ(START_SVC)

' Clean Up
Set objWMI = Nothing

' Exit script
WScript.Quit(0)


'//////////////////////////////////////////////////////////////////////////////
'// METHODS / FUNCTIONS

'// StartServ: Starts a given service name (if it exists).
Sub StartServ (sServiceName)
	Dim colServices 
	Dim objService

	' Get Service collection
	Set colServices = objWMI.ExecQuery ("SELECT * FROM Win32_Service WHERE Name ='" & sServiceName & "'") 
	If Not (colServices Is Nothing) Then

		' Start Service - should be only one item in the Services collection
		For Each objService in colServices
			objService.StartService()
		Next 
					
	End If

	' Clean up
	Set colServices = Nothing
	Set objService = Nothing
End Sub

'// new section
'Start Service
'Not used - left as and example of the StartService that assumes service is really stopped
strServiceName = "Alerter"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
    objService.StartService()
Next

'// RestartServ: Same as StartServ, just Stops first, waits 10 seconds, then Starts
Sub RestartServ(sServiceName)
	Dim colServices 
	Dim objService
	On Error Resume Next
	
	' Get Service collection
	Set colServices = objWMI.ExecQuery ("SELECT * FROM Win32_Service WHERE Name ='" & sServiceName & "'") 
	If Not (colServices Is Nothing) Then

		' Stop service first
		For Each objService in colServices			
			objService.StopService()
			'Wscript.Echo "Tried to stop " & objService.Name
		Next 
		
		' Wait 10 seconds
		wScript.Sleep 120000
		
		' Start service
		For each objService In colServices			
			objService.StartService()
			'Wscript.Echo "Tried to start..."
		Next
		
	End If

	' Clean up
	Set colServices = Nothing
	Set objService = Nothing
End Sub