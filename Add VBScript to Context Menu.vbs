Option Explicit

Dim oFSO, oFile, oShell, oWMI
Dim colProcList, oProc

Set oFSO = CreateObject("Scripting.FileSystemObject")
If oFSO.FileExists("C:\Windows\ShellNew\MyScript.vbs") Then
	'do nothing
Else
	Set oFile = oFSO.CreateTextFile("C:\Windows\ShellNew\MyScript.vbs")
End If

Set oShell = CreateObject("WScript.Shell")
oShell.RegWrite "HKCR\.VBS\ShellNew\FileName","MyScript.vbs"

Set oWMI = GetObject("WinMgmts://")
Set colProcList = oWMI.ExecQuery _
    ("SELECT * FROM Win32_Process WHERE Name = 'explorer.exe'")
For Each oProc in colProcList
    oProc.Terminate()
Next

WScript.Echo "Done"