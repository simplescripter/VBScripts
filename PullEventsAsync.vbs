Option Explicit

Dim oFSO, oFile, sComputerFile, sComputer
Dim SINK, oWMI
Const ForReading = 1

On Error Resume Next

sComputerFile = "C:\Script\Clients.txt"
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.OpenTextFile(sComputerFile, ForReading)
Set SINK = WScript.CreateObject("WbemScripting.SWbemSink","SINK_")

Do Until oFile.AtEndOfStream
	sComputer = oFile.ReadLine
	Set oWMI = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\" _
		& sComputer & "\root\cimv2")
	retValue = oWMI.ExecQueryAsync(SINK, "Select * From Win32_nTLogEvent" _
		& " Where LogFile = 'System'")
Loop

MsgBox "Pause here"

Sub SINK_OnObjectReady(oObject, oAsyncContext)
	WScript.Echo oObject.EventCode & "  " & oObject.ComputerName
End Sub

Sub SINK_OnCompleted(i,j,k)

End Sub
	