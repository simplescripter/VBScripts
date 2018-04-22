'******************************************************************************
'
' This script uses an asynchronous WMI query to monitor for specific logons
' across a set of systems.  Adapted from a script found at the ScriptCenter
' at Microsoft.com/technet/scriptcenter
'
' Shawn Stugart
' 11-18-2005
'
'******************************************************************************

Option Explicit

Dim arrComputers(), oFSO, oInputFile, sInputFile
Dim sComputer, oWMI, colSystems, oSystem
Dim sUserName, sAsyncQuery, SINK, intSize
Dim hostname, iHelp, oShell, sQuery, colLogons
Dim oLogon, sError, iError

CheckCScript

sUserName = "Classnet\Administrator" 'Name of account to monitor
sInputFile = "C:\Clients.txt" 'List of systems to monitor

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oInputFile = oFSO.OpenTextFile(sInputFile,1)
Set SINK = WScript.CreateObject("WbemScripting.SWbemSink","SINK_")

intSize = 0
Do Until oInputFile.AtEndOfStream
	Redim Preserve arrComputers(intSize)
	arrComputers(intSize) = oInputFile.ReadLine
	sComputer = arrComputers(intSize)
  	WScript.Echo vbCrLf & "Host: " & UCase(sComputer) & vbCrLf
    iError = ShowLogons(sComputer)
    If iError = 0 Then
    	TrapLogons(sComputer)
    End If
  	intSize = intSize + 1
Loop

WScript.Echo vbCrLf & vbCrLf & "***************  MONITORING MODE  ***************" & vbCrLf

Do
	WScript.Sleep 1000
Loop

'***************************************************************************
Function ShowLogons(sHost)
	
	On Error Resume Next
	
	sQuery = "SELECT * FROM Win32_ComputerSystem"
	Set oWMI = GetObject("winmgmts:" _
 		& "{impersonationLevel=impersonate}!\\" & sHost & "\root\cimv2")
 	If Err = 0 Then
 		Set colLogons = oWMI.ExecQuery(sQuery)
 		For each oLogon in colLogons
 			If LCase(oLogon.UserName) = LCase(sUserName) Then
 			 	WScript.Echo vbTab & sUserName & " is currently logged " _
 					& "on to " & oLogon.Name & "!"
 			End If
 		Next
 		ShowLogons = 0
 	Else
 		HandleError(sComputer)
 		WScript.Echo "  Unable to monitor logons on " & sComputer
 		ShowLogons = 1
 	End If
 
 End Function		

'***************************************************************************
Sub TrapLogons(sHost)
	
	On Error Resume Next
	
	sAsyncQuery = "SELECT * FROM __InstanceModificationEvent WITHIN 1 " & _
	 "WHERE TargetInstance ISA 'Win32_ComputerSystem'"
	Set oWMI = GetObject("winmgmts:" _
	 	& "{impersonationLevel=impersonate}!\\" & sHost & "\root\cimv2")
	If Err = 0 Then
	  	oWMI.ExecNotificationQueryAsync SINK, sAsyncQuery
	  	If Err = 0 Then
	    	WScript.Echo vbCrLf
		Else
	    	HandleError(sHost)
	    	WScript.Echo "  Unable to monitor logons."
		End If
	Else
	  	HandleError(sHost)
	  	WScript.Echo "  Unable to monitor logons."
	End If

End Sub

'******************************************************************************

Sub SINK_OnObjectReady(objLatestEvent, objAsyncContext)

  If LCase(objLatestEvent.TargetInstance.UserName) = LCase(sUserName) Then
    Wscript.Echo VbCrLf & "User: " & objLatestEvent.TargetInstance.UserName
    Wscript.Echo "  Logged On To: " & objLatestEvent.TargetInstance.Name
    Wscript.Echo "  Time: " & Now
  End If

End Sub

'******************************************************************************
Sub HandleError(sHost)

	sError = VbCrLf & "  ERROR on " & sHost & VbCrLf & _
	 "  Number: " & Err.Number & VbCrLf & _
	 "  Description: " & Err.Description & VbCrLf & _
	 "  Source: " & Err.Source
	WScript.Echo sError
	Err.Clear

End Sub

'******************************************************************************
Sub CheckCScript()
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    iHelp = MsgBox("At a command line, run " & Chr(34) _
	    	& "cscript " & WScript.ScriptFullName & Chr(34) _
	    	& vbCrLf & vbCrLf & "Do you want help?", vbYesNo)
	    If iHelp = vbYes Then
	    	Set oShell = CreateObject("WScript.Shell")
	    	oShell.Run("cmd.exe")
	    	WScript.Sleep 500
	    	oShell.SendKeys "cscript.exe " & Chr(34) & WScript.ScriptFullName & Chr(34)
	    	WScript.Quit
	    Else
	    	WScript.Quit
	    End If
	End If
End Sub