const HKEY_LOCAL_MACHINE = &H80000002

strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion" ' Registry path
strValueName = "RegisteredOrganization" ' Registry value

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
errReturn = oReg.GetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
If errReturn = 0 Then
	iInput = MsgBox(strValueName & " = " & strValue & vbCrLf _
		& vbCrLf & "CHANGE VALUE?", vbYesNo)
	If iInput = vbYes Then
		sNewValue = Inputbox("Enter New String:")
		errReturn = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,sNewValue)
		WScript.Echo "Changed " & strValueName _
			& " to " & sNewValue
	Else		
		WScript.Quit
	End If
Else
	WScript.Echo "ERROR."
End If