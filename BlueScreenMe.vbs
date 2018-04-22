'------------------------------------------------------------------------
'  Adds the CrashOnCtrlScroll value to the target registry.  After a
'  reboot, you can intentionally blue screen the target system using
'  the following keystroke combination: Ctrl+Scroll,Scroll
'
'
'  Shawn Stugart
'========================================================================
Option Explicit

Dim sKeyPath1, sKeyPath2, sValueName, errReturn1
Dim sComputer, oReg, sValue, errReturn2

Const HKEY_LOCAL_MACHINE = &H80000002
const ForReading = 1
sKeyPath1 = "SYSTEM\CurrentControlSet\Services\kbdhid\Parameters\"
sKeyPath2 = "SYSTEM\CurrentControlSet\Services\i8042prt\Parameters\"
sValueName = "CrashOnCtrlScroll"
sValue = "1"

On Error Resume Next

sComputer = InputBox("Enter a computer to prep for Blue Screen:")
If sComputer = "" Then WScript.Quit
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
	sComputer & "\root\default:StdRegProv")
If Err.Number <> 0 Then
     sResult = sResult & "An error has occurred on " & sComputer _
     	& ".  You may have mistyped the computer name." & VbCrLf
     Err.Clear
End If
errReturn1 = oReg.SetDWORDValue(HKEY_LOCAL_MACHINE,sKeyPath1,sValueName,sValue)
errReturn2 = oReg.SetDWORDValue(HKEY_LOCAL_MACHINE,sKeyPath2,sValueName,sValue)

WScript.Echo "Finished"