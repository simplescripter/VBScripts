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
	    errReturnCode1 = oService.StopService()
	    errReturnCode2 = oService.Change( , , , , "Disabled")   
	Next
	If errReturnCode1 = 0 Then
		WScript.Echo "Stopped Automatic Updates on " & sComputer
	Else
		WScript.Echo "FAILED on " & sComputer
	End If
	If errReturnCode2 = 0 Then
		WScript.Echo "Disabled Automatic Updates on " & sComputer
	Else
		WScript.Echo "FAILED on " & sComputer
	End If
	WScript.Echo "--------------------------------------------------------"
Loop