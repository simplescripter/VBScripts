sComputer = "London"
const HKEY_CLASSES_ROOT = &H80000000
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
	sComputer & "\root\default:StdRegProv")
If Err.Number <> 0 Then
    MsgBox "An error has occurred.  You may have mistyped the computer name."
End If

sValueName = ""
sValue1 = "Open in Command Shell"
sValue2 = "cmd.exe /k cd %1"
sKeyPath1 = "Directory\Shell\OpenNew"
sKeyPath2 = "Directory\Shell\OpenNew\Command"
Return = oReg.CreateKey(HKEY_CLASSES_ROOT, sKeyPath1) 
If Return = 0 Then
	Return = oReg.CreateKey(HKEY_CLASSES_ROOT, sKeyPath2)
	If Return = 0 Then
		Return1 = oReg.SetStringValue( _
        	HKEY_CLASSES_ROOT,sKeyPath1,sValueName,sValue1)
        Return2 = oReg.SetStringValue( _
        	HKEY_CLASSES_ROOT,sKeyPath2,sValueName,sValue2)
		If (Return1 = 0) AND (Return2 = 0) AND (Err.Number) = 0 Then
			MsgBox "Success"
		End If
	Else
		MsgBox "Failed"
	End If
Else
	MsgBox "Failed"
End If
