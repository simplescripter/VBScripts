Option Explicit

Const ForReading = 1

Dim sComputer, oFSO, sResult
Dim oWMI, colAdapters, oAdapter
Dim oFile

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.OpenTextFile("C:\Clients.txt", ForReading)

On Error Resume Next

Do Until oFile.AtEndOfStream
	sComputer = oFile.ReadLine
	Set oWMI = GetObject("winmgmts:" & "!\\" & sComputer & "\root\cimv2")
	Set colAdapters = oWMI.ExecQuery("Select * from " _
	    & "Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	For Each oAdapter in colAdapters
		oAdapter.RenewDHCPLease
	Next
	If Err = 0 Then
		sResult = sResult & "Renewed IP Settings on " & sComputer & VbCrLf
	Else
		sResult = sResult & "FAILED on " & sComputer & VbCrLf
		Err.Clear
	End If
	WScript.Echo sResult
	sResult = ""
Loop

