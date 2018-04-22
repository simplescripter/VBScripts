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
	Dim iPingTimeOut, oShell, oScriptExec, sPingStdOut
	iPingTimeOut = "100"
	Set oShell = CreateObject("WScript.Shell")
	Set oScriptExec = oShell.Exec("ping -n 1 -w " & iPingTimeOut _
		 & " " & sTarget)
    sPingStdOut = Lcase(oScriptExec.StdOut.ReadAll)
    If InStr(sPingStdOut, "reply from") Then
		Pinger = 1
	Else
	    Pinger = 0
	End If
End Function