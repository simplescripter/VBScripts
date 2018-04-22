'------------------------------------------------------------------------
'
'  An example script using WMI to set a registry value on a number of
'  hosts as defined in the text file "c:\clients.txt".  This particular
'  registry edit prompts you to change the AutoAdminLogon value.
'
'  Shawn Stugart
'========================================================================
Option Explicit

Dim strKeyPath, strValueName, sFileName, oFSO, oShell, oInputFile
Dim strComputer, oReg, sResult, vChangeValue, strValue, errReturn

Const HKEY_LOCAL_MACHINE = &H80000002
const ForReading = 1
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
strValueName = "AutoAdminLogon"
sFileName = "clients.txt"

On Error Resume Next

set oFSO = CreateObject("Scripting.FileSystemObject")
set oShell = CreateObject("WScript.Shell")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)
Do Until oInputFile.AtEndOfStream
    strComputer = oInputFile.Readline
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
    strComputer & "\root\default:StdRegProv")
    If Err.Number <> 0 Then
         sResult = sResult & "An error has occurred on " & strComputer _
         	& ".  You may have mistyped the computer name." & VbCrLf
         Err.Clear
    End If
    oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
    sResult = sResult & strComputer & ": " & strValue & VbCrLf
Loop

vChangeValue = MsgBox(sResult & VbCrLf & VbCrLf & "Do you want to change the " _
	& "AutoAdminLogon values?", vbYesNo)
sResult = ""
If vChangeValue = vbYes Then
	strValue = InputBox("What value do you want to set the AutoAdminLogon " _
    	& "to?","0 or 1?","0")
    oInputFile.Close
    Set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)
    Do Until oInputFile.AtEndOfStream
	    strComputer = oInputFile.Readline
	    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
	    	strComputer & "\root\default:StdRegProv")
	    If Err.Number <> 0 Then
	         sResult = sResult & "An error has occurred on " & strComputer _
	         	& ".  You may have mistyped the computer name." & VbCrLf
	         Err.Clear
	    Else
	    	errReturn = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
	    End If
	Loop
End If
WScript.Echo "Finished"