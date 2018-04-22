'------------------------------------------------------------------------
'
'  An example script using WMI to set a registry value on a number of
'  hosts as defined in the text file "c:\clients.txt".  This particular
'  registry edit prompts you to change the Trust Policy value, which 
'  controls whether or not clients will execute scripts that have
'  not been digitally signed.
'
'  Shawn Stugart
'========================================================================

const HKEY_LOCAL_MACHINE = &H80000002
const ForReading = 1

On Error Resume Next

sFileName = "clients.txt"
set oFSO = CreateObject("Scripting.FileSystemObject")
set oShell = CreateObject("WScript.Shell")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)
dwValue = InputBox("What value do you want to set the Trustpolicy " _
    & "to?","0, 1, or 2?","0")
Do Until oInputFile.AtEndOfStream
    strComputer = oInputFile.Readline
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
    strComputer & "\root\default:StdRegProv")
    If Err.Number <> 0 Then
         WScript.Echo "An error has occurred.  You may have mistyped the computer name."
         WScript.Quit
    End If
strKeyPath = "SOFTWARE\Microsoft\Windows Script Host\Settings"
strValueName = "TrustPolicy"
oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
WScript.Echo "Trust Policy set to " & dwValue & " on " & strComputer
Loop
     

