const HKEY_LOCAL_MACHINE = &H80000002
strComputer = InputBox("Enter the MACHINE NAME on which you would like to enable Remote Desktop:")
If strComputer = "" Then
  WScript.Quit
End If
On Error Resume Next
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")
If Err.Number <> 0 Then
  WScript.Echo "An error has occurred.  You may have mistyped the computer name."
  WScript.Quit
End If
strKeyPath = "SYSTEM\CurrentControlSet\Control\Terminal Server"
strValueName = "fDenyTSConnections"
oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
If dwValue = 1 Then
  prompt = MsgBox ("Remote Desktop is Currently disabled.  Do you want to ENABLE it?", vbYesNo)
  If prompt = vbYes then
    dwValue = 0
    oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
     WScript.Echo "Remote Desktop is now ENABLED on " & strComputer
    WScript.Quit
  ElseIf prompt = vbNo then
    WScript.Echo "Remote Desktop is still DISABLED."
  	Wscript.Quit
  End If
ElseIf dwValue = 0 then
  prompt = MsgBox ("Remote Desktop is Currently ENABLED.  Do you want to DISABLE it?", vbYesNo)
  If prompt = vbYes then
    dwValue = 1
    oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
    WScript.Echo "Remote Desktop is now DISABLED on " & strComputer
    WScript.Quit
  ElseIf prompt = vbNo then
    WScript.Echo "Remote Desktop is still ENABLED."
    WScript.Quit
  End If
End If

