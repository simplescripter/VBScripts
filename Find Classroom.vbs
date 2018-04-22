'------------------------------------------------------------------------
'
'  
'
'  Shawn Stugart
'========================================================================
Option Explicit

const ForReading = 1

Dim sFileName, oFSO, oFSO2, oShell, oInputFile, dwValue, oOutputFile
Dim strComputer, oReg, strKeyPath, strValueName1, strValueName2
Dim strResult, oNet, strDesktop, oURLLink, strSetup, oWMI

On Error Resume Next

sFileName = "clients.txt"
set oFSO = CreateObject("Scripting.FileSystemObject")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)
Set oOutputFile = oFSO.CreateTextFile("c:\output.txt",True)

Do Until oInputFile.AtEndOfStream
    strComputer = oInputFile.Readline
    WScript.Echo "Trying " & strComputer & vbCrLf
    Set oWMI = GetObject("Winmgmts:\\" & strComputer)
    If Err.Number = 0 Then
        oOutputFile.WriteLine(strComputer)
    Else
        Err.Clear
    End If
Loop

     

