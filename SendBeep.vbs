sComputer = InputBox("Machine to Beep")

On Error Resume Next

WScript.Sleep 5000
SendBeep(sComputer)
WScript.Sleep 500
SendBeep(sComputer)
WScript.Sleep 400
SendBeep(sComputer)
WScript.Sleep 200
SendBeep(sComputer)
WScript.Sleep 500
SendBeep(sComputer)
WScript.Sleep 1000
SendBeep(sComputer)
WScript.Sleep 500
SendBeep(sComputer)

WScript.Echo "Done beeping."

Sub SendBeep (sBoxName)
  Set Proc = GetObject("WinMgmts:\\" & sBoxName & "\root\cimv2:Win32_Process")
  If Err <> 0 Then
  	WScript.Echo "Error beeping " & sBoxName
  	WScript.Quit
  End If
  errReturn = Proc.Create("Rundll32.exe user32.dll,MessageBeep", null, null, intProcessID)
End Sub