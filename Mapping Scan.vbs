const ForReading = 1
sFileName = "c:\clients.txt"
sUser = "Administrator"
sPassword = "password"
set oFSO = CreateObject("Scripting.FileSystemObject")
set oShell = CreateObject("WScript.Shell")
set oInputFile = oFSO.OpenTextFile(sFileName, ForReading)
Set oNet = CreateObject("WScript.Network")

On Error Resume Next

oNet.RemoveNetworkDrive "X:",True
Err.Clear

Do Until oInputFile.AtEndOfStream
    sCompName = oInputFile.Readline
    oNet.MapNetworkDrive "X:", "\\" & sCompName & "\C$",, sUser, sPassword
    If Err.Number = 0 Then
    	WScript.Echo "***SUCCESS***" & vbTab & "Mapped as " & sUser & " to " & sCompName
    	oNet.RemoveNetworkDrive "X:", True
    Else
    	WScript.Echo sCompName & ": FAILED mapping as " & sUser
    	Err.Clear
    End If
Loop