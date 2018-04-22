Option Explicit

Dim strComputer, oReg, strKeyPath, strValueName
Dim dwValue, valPrompt, errReturn 

Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = InputBox("Enter Computer Name:")
If strComputer = "" Then
  WScript.Quit
End If
'On Error Resume Next
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
	strComputer & "\root\default:StdRegProv")
strKeyPath = "SYSTEM\CurrentControlSet\Services\USBSTOR"
strValueName = "Start"
errReturn = oReg.GetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue)
If errReturn <> 0 Then
	MsgBox "USB Storage has not been used on " & strComputer
	WScript.Quit
End If
If Err = 0 Then
	If dwValue = 4 Then
	    valPrompt = MsgBox("USB Storage is currently disabled.  Do you want to ENABLE it?", vbYesNo)
	    If valPrompt = vbYes Then
		    dwValue = 3
		    oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
		    MsgBox "USB Storage is now ENABLED on " & strComputer
	    ElseIf valPrompt = vbNo Then
		    MsgBox "USB Storage is still DISABLED."
	    End If
	ElseIf dwValue = 3 Then
	  	valPrompt = MsgBox ("USB Storage is Currently ENABLED.  Do you want to DISABLE it?", vbYesNo)
	  	If valPrompt = vbYes Then
		    dwValue = 4
		    oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
		    MsgBox "USB Storage is now DISABLED on " & strComputer
	    ElseIf valPrompt = vbNo Then
		    MsgBox "USB Storage is still ENABLED on " & strComputer
	    End If
	Else
		MsgBox "Couldn't set USB Storage on " & strComputer & vbCrLf _
		& "Check that the remote system is available."
	End If
Else
	MsgBox "Couldn't set USB Storage on " & strComputer & vbCrLf _
		& "Check that the remote system is available."	
End If
