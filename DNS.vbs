Option Explicit

Const ForReading = 1
dim strComputer, strDNSServer, strFileName, objFSO
dim objWMIService, colAdapters, objAdapter
Dim objShell, objInputFile


strFileName = "C:\clients.txt"
strDNSServer = Array("192.168.8.200","192.168.25.1")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Set objInputFile = objFSO.OpenTextFile(strFileName, ForReading)

Do Until objInputFile.AtEndOfStream
	strComputer = objInputFile.Readline
	Set objWMIService = GetObject("winmgmts:" & "!\\" & strComputer & "\root\cimv2")
	Set colAdapters = objWMIService.ExecQuery("Select * from " _
	    & "Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	For Each objAdapter in colAdapters
		objAdapter.SetDNSServerSearchOrder strDNSServer
	Next
Loop
WScript.Echo "DONE"

