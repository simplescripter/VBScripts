const ForReading = 1
sFileName = "C:\Clients.txt"
set oFSO = CreateObject("Scripting.FileSystemObject")
set oInputFile = oFSO.OpenTextFile(sFileName, ForReading)

On Error Resume Next

Do Until oInputFile.AtEndOfStream
	sComputer = oInputFile.ReadLine
	Set oWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
	Set colServiceList = oWMIService.ExecQuery _
	        ("Select * from Win32_Service where Name='wuauserv'")
	For Each oService in colServiceList
	    WScript.Echo sComputer & "	" & oService.State
		WScript.Echo "--------------------------------------------------------"
	Next
Loop