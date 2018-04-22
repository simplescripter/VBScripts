Option Explicit
On Error Resume Next
Dim strKeyPath, strComputer, objReg, arrSubKeys, key
Const HKEY_LOCAL_MACHINE = &H80000002

strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
strComputer = "."

Set objReg=GetObject("WinMgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
If Err.Number <> 0 Then
	WScript.Echo "Couldn't read from the registry."
	Err.Clear
Else
	objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
	For Each key in arrSubKeys
		Select Case key
			Case "{3E713D52-C967-41FB-AA24-3A92CC1025A4}"
				WScript.Echo "Found Remote Desktop Connection"
			Case "{6F716D8C-398F-11D3-85E1-005004838609}"
				WScript.Echo "Found Webfldrs"
			Case "{8A11A031-928C-471B-A6C5-4A1C99607573}"
				WScript.Echo "Found PrimalScript 4"
		End Select
	Next
End If