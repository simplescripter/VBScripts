sComputer = InputBox("Enter Computer Name:")
If Pinger(sComputer) = 1 Then
	MyCode(sComputer)
Else
	WScript.Echo sComputer & " is not responding."
	WScript.Quit
End If

Sub MyCode(sSystem)
	Set oWMI = GetObject("WinMgmts://" & sSystem)
	Set col = oWMI.InstancesOf("Win32_Process")
	For Each obj In col
		sResult = sResult & obj.Name & VbCrLf
	Next
	WScript.Echo sResult
End Sub

Function Pinger(sTarget)
	Dim iPingTimeOut, colItems, oItem, oWMI 
	iPingTimeOut = "100"
	On Error Resume Next
	Set oWMI = GetObject("WinMgmts://")
	Set colItems = oWMI.ExecQuery("Select * from Win32_PingStatus " _
		& "Where Address = '" & sTarget & "'" _
		& "AND Timeout =" & iPingTimeOut, "WQL", 48)
	If Err.Number <> 0 Then
		WScript.Echo "An error has occurred.  This script " _
			& "must be run on Windows XP/Windows Server 2003 " _
			& "or later."
		WScript.Quit
	End If
	For Each oItem In colItems
	   If oItem.StatusCode = 0 Then
	       Pinger = 1
	   Else
	   	   Pinger = 0
	   End If
	Next
End Function