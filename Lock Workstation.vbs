sComputer = InputBox("Machine to Lock")

LockBox "Rundll32.exe user32.dll,LockWorkStation", sComputer

Sub LockBox (sCommandLine, sBoxName)
  Set Proc = GetObject("WinMgmts:\\" & sBoxName & "\root\cimv2:Win32_Process")
  errReturn = Proc.Create(sCommandLine, null, null, intProcessID)
  If errReturn = 0 Then
	MsgBox sBoxName & " has been locked."
  Else
	MsgBox "Failed to lock " & sBoxName
  End If
  Set Proc = Nothing
End Sub