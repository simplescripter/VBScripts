'---------------------------------------------------------------------------
'
'   Shawn Stugart
'===========================================================================


Option Explicit

dim strKeyPath, sFileName, oFSO, oShell, oInputFile
dim strComputer, objReg, arrSubKeys, subkey, sResult
dim key, j, objWMIService, colItems, objItem, strUserName
dim strEFSUsers, strNeverUsed, sKey, strValueName, dwValue

Const HKEY_LOCAL_MACHINE = &H80000002
const ForReading = 1

strValueName ="Start"
sFileName = "clients.txt"
sKey = "USBSTOR"

set oFSO = CreateObject("Scripting.FileSystemObject")
set oShell = CreateObject("WScript.Shell")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)

On Error Resume Next

Do Until oInputFile.AtEndOfStream
    strComputer = oInputFile.Readline
    strKeyPath = "SYSTEM\CurrentControlSet\Services"
    Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\default:StdRegProv")
    If Err.Number <> 0 Then
        sResult = sResult & "Couldn't read " & strComputer & vbCrLf
        Err.Clear
    Else
	    objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
	    j = 0
	    For Each key in arrSubkeys
	        If key = sKey Then
	            j = 1
	            Exit For
	        End If
	    Next
	    If j = 1 Then
	    	WScript.Echo "USB Storage has been used on " & strComputer
	    Else
	    	WScript.Echo strComputer & ": CLEAN"
	    	'CreateKey
	    End If
	End If
Loop

Sub CreateKey()
	strKeyPath = strKeyPath & "\" & sKey
	objReg.CreateKey HKEY_LOCAL_MACHINE, strKeyPath
	dwValue = 4
	If Err = 0 Then
		WScript.Echo "set"
	Else
		WScript.Echo "error"
	End If
	objReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName
End Sub