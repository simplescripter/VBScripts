const HKEY_LOCAL_MACHINE = &H80000002

strKeyPath = "SOFTWARE\Microsoft\Test" ' Registry path
strValueName = "SMSThing" ' Registry value

strComputer = InputBox("Enter the MACHINE NAME:")
If strComputer = "" Then
  WScript.Quit
End If
On Error Resume Next
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")
If Err.Number <> 0 Then
  WScript.Echo "An error has occurred.  You may have mistyped " _
  & "the computer name, or the remote system may not be available."
  WScript.Quit
End If
errReturn = oReg.GetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue)
If errReturn = 0 Then
	If dwValue = 1 Then
	  prompt = MsgBox ("The value is currently set to 1.  Change to 0?", vbYesNo)
	  If prompt = vbYes Then
	    dwValue = 0
	    oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	    WScript.Quit
	  ElseIf prompt = vbNo Then
	  	Wscript.Quit
	  End If
	ElseIf dwValue = 0 Then
	  prompt = MsgBox ("The value is currently set to 0.  Change to 1?", vbYesNo)
	  If prompt = vbYes Then
	    dwValue = 1
	    oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	    WScript.Quit
	  ElseIf prompt = vbNo Then
	    WScript.Quit
	  End If
	End If
Else
	MsgBox "Error reading registry value.  The value " _
		& "or registry key may not exist.", , "ERROR"
	WScript.Quit
End If

