'---------------------------------------------------------------------------
'
'   This script loads a text file called c:\clients.txt and scans the
'   computers listed to find out if a particular string entry exists
'	in the registry.
'
'   Shawn Stugart
'===========================================================================


Option Explicit

Dim sKeyPath, sFileName, oFSO, oShell, oInputFile
Dim sComputer, oReg, sResult
Dim sValueName, sRegValue, sReturn

Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
const ForReading = 1

sKeyPath = "SOFTWARE\Adobe\CommonFiles"
sValueName = "RegTestValue"
sRegValue = "{0002443B-0000-0000-C000-000000000046}"
sFileName = "clients.txt"

set oFSO = CreateObject("Scripting.FileSystemObject")
set oShell = CreateObject("WScript.Shell")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)

On Error Resume Next

Do Until oInputFile.AtEndOfStream
    sComputer = oInputFile.Readline
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
        sComputer & "\root\default:StdRegProv")
    If Err.Number <> 0 Then
        sResult = sResult & "Couldn't read " & sComputer & vbCrLf
        Err.Clear
    Else
	    sReturn = oReg.GetStringValue(HKEY_LOCAL_MACHINE, sKeyPath, sValueName, sRegValue)
	    If sReturn = 0 Then
	    	sResult = sResult & "String found on " & sComputer & vbCrLf
	    Else
	    	sResult = sResult & "Couldn't find string on " & sComputer & vbCrLf
	    End If
	End If 
Loop

WScript.Echo sResult