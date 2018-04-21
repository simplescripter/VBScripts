'------------------------------------------------------------------------
'
'  
'
'  Shawn Stugart
'========================================================================
Option Explicit
const HKEY_LOCAL_MACHINE = &H80000002
const ForReading = 1

Dim sFileName, oFSO, oShell, oInputFile, oNet, oURLLink2
Dim strComputer, strResult, strDesktop, oURLLink1, strSetup

On Error Resume Next

sFileName = "clients.txt"
set oFSO = CreateObject("Scripting.FileSystemObject")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)
Do Until oInputFile.AtEndOfStream
    strComputer = oInputFile.Readline
    set oNet = CreateObject("WScript.Network")
    oNet.MapNetworkDrive "X:", "\\" & strComputer & "\C$"
	strDesktop = "\Documents and Settings\All Users\Desktop"
	set oShell = CreateObject("WScript.Shell")
	set oURLLink1 = oShell.CreateShortcut("X:" & strDesktop & "\2266 Course Evaluation.url")
	oURLLink1.TargetPath = "http://www.metricsthatmatter.com/usnhco/spr"
	oURLLink1.Save
        set oURLLink2 = oShell.CreateShortcut("X:" & strDesktop & "\Integrated Learning Manager.url")
	oURLLink2.TargetPath = "http://my.newhorizons.com"
	oURLLink2.Save
	oNet.RemoveNetworkDrive "X:", True
	If Err.Number = 0 Then
	    WScript.Echo "Shortcut copied to " & strComputer
	Else
	    WScript.Echo "FAILED. Could not copy shortcut to " & strComputer
	    Err.Clear
	End If
Loop

     
