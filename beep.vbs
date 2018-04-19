
For i = 1 To 3
	Set oShell = CreateObject("WScript.Shell")
	oShell.Run("RunDll32.exe user32.dll,MessageBeep")
	WScript.Sleep 1000
Next