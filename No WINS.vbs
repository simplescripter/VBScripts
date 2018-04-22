Option Explicit
On Error Resume Next
dim sFileName, oFSO, oInputFile, strComputer
dim objWMIService, colAdapters, objAdapter, hostName
const ForReading = 1
sFileName = "clients.txt"

Call CheckCScript

set oFSO = CreateObject("Scripting.FileSystemObject")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)
If Err.Number <> 0 Then
    WScript.Echo "Couldn't Find C:\" & sFileName
    WScript.Echo Err.Number
    WScript.Quit
End If

Do Until oInputFile.AtEndOfStream
    strComputer = oInputFile.Readline
	Set objWMIService = GetObject("winmgmts:" & "!\\" & strComputer & "\root\cimv2")
	If Err <> 0 Then
	    Err.Clear
	    WScript.Echo "Could not contact " & strComputer
	Else
		Set colAdapters = objWMIService.ExecQuery("Select * from " _
		    & "Win32_NetworkAdapterConfiguration Where IPEnabled = True")
		For Each objAdapter in colAdapters
			objAdapter.SetWinsServer "",""
		Next
		WScript.Echo strComputer & " done."
    End If
Loop

Sub CheckCScript
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    WScript.Quit
	End If
End Sub