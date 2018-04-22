Option Explicit

Dim sRunMRUPath, sValueName, sComputer, oReg, errReturn
Dim oWMI, colProcList, oProc, sResult

Const HKEY_CURRENT_USER = &H80000001
sRunMRUPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
sValueName = "MRUList"

sComputer = InputBox("Enter the MACHINE NAME:")
If sComputer = "" Then
  WScript.Quit
End If
On Error Resume Next
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
	sComputer & "\root\default:StdRegProv")
If Err.Number <> 0 Then
  sResult = sResult & "An error has occurred on " & sComputer & "." & VbCrLf
  Err.Clear
Else
	errReturn = oReg.DeleteValue(HKEY_CURRENT_USER,sRunMRUPath,sValueName)
	If errReturn = 0 Then
		Set oWMI = GetObject("WinMgmts://" & sComputer)
		Set colProcList = oWMI.ExecQuery _
	    ("SELECT * FROM Win32_Process WHERE Name = 'explorer.exe'")
		For Each oProc in colProcList
	    	oProc.Terminate()
		Next
	Else		
		sResult = sResult & "Error deleting value on " & sComputer & ". The value may be empty already." & VbCrLf
		Err.Clear
	End If
End If

If sResult <> "" Then
	WScript.Echo sResult
Else
	WScript.Echo "Complete."
End If
