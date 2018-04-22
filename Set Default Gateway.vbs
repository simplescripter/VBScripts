Option Explicit
dim strComputer, strGateway, strFileName, objFSO
dim objWMIService, colAdapters, objAdapter, objInputFile
const ForReading = 1

strFileName = "clients.txt"
set objFSO = CreateObject("Scripting.FileSystemObject")
set objInputFile = objFSO.OpenTextFile("c:\" & strFileName, ForReading)

Do Until objInputFile.AtEndOfStream
	strComputer = objInputFile.ReadLine
	strGateway = Array("192.168.1.200")
	Set objWMIService = GetObject("winmgmts:" & "!\\" & strComputer & "\root\cimv2")
	Set colAdapters = objWMIService.ExecQuery("Select * from " _
	    & "Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	For Each objAdapter in colAdapters
		objAdapter.SetGateways strGateway
	Next
	If Err.Number <> 0 Then
		WScript.Echo "Error setting gateway on " & strComputer
	Else
		WScript.Echo "DONE setting gateway on " & strComputer
	End If
Loop

