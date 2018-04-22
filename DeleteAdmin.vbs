' Shawn Stugart
'
' For Casey at LeaderQuest
'
' 6-16-2010

Option Explicit

Dim sClients, oFSO, oFile, oUser
Dim sResults, sAdmin, sComputer
Const ForReading = 1

On Error Resume Next

sClients = "C:\Script\Clients.txt"
sAdmin = InputBox("Enter the name of the local administrator account to delete:", "Admin Name")
If sAdmin = "" Then WScript.Quit
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.OpenTextFile(sClients, ForReading)
If Err <> 0 Then
	WScript.Echo "Error reading " & sClients & ". Make sure the file " _
		& "exists and contains your list of clients."
	WScript.Quit
End If

Do Until oFile.AtEndOfStream
	sComputer = oFile.ReadLine
	Set oUser = GetObject("WinNT://" & sComputer)
	If Err <> 0 Then
		sResults = sResults & "Error binding to " & sComputer & VbCrLf
		Err.Clear
	End If
	oUser.delete "User", sAdmin
	If Err <> 0 Then
		sResults = sResults & "Error deleting " & sAdmin & " on " & sComputer & VbCrLf
		Err.Clear
	Else
		sResults = sResults & sAdmin & " successfully deleted from " & sComputer & VbCrLf
	End If
Loop

WScript.Echo sResults