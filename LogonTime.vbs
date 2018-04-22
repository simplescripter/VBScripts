'******************************************************************************
'
' This script was designed for student logons in the New Horizons 
' Mentored Learning Center.  It creates a simple text file called 
' LogonTime in the user's profile and adds a time stamp that can 
' be harvested by ProfileScan2.vbs.
'
' Shawn Stugart
' 9-6-2006
'
'******************************************************************************

Option Explicit

Dim oShell, oEnv, oFSO, sUserLog, oFile

Const ForWriting = 2

Set oShell = CreateObject("WScript.Shell")
Set oEnv = oShell.Environment("Process")
Set oFSO = CreateObject("Scripting.FileSystemObject")
sUserLog = oEnv("USERPROFILE") & "\LogonTime"
If oFSO.FileExists(sUserLog) Then
	Set oFile = oFSO.OpenTextFile(sUserLog, ForWriting)
	oFile.WriteLine(Now)
	oFile.Close
Else
	Set oFile = oFSO.CreateTextFile(sUserLog)
	oFile.WriteLine(Now)
	oFile.Close
End If
	
