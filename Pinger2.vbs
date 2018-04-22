Option Explicit

Dim oWMI, sSubnet, i, sHost
Dim sTarget, colItems, oItem, iTimeout

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

On Error Resume Next

iTimeout = 25
sSubnet = "192.168.1."
Set oWMI = GetObject("winmgmts:\\")

For i = 100 To 150
	sHost = sSubnet & i
	Pinger(sHost)
Next

Sub Pinger(sTarget)
	Set colItems = oWMI.ExecQuery("Select * from Win32_PingStatus Where Address = '" & sTarget & "'" _
		& "AND Timeout =" & iTimeout, "WQL", _
		wbemFlagReturnImmediately + wbemFlagForwardOnly)
	For Each oItem In colItems
	   If oItem.StatusCode = 0 Then
	       WScript.Echo sTarget & " --- ALIVE"
	   Else
	   	   WScript.Echo "XX " & sTarget & " XX"
	   End If
	Next
End Sub