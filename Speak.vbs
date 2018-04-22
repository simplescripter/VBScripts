Option Explicit

'On Error Resume Next

Dim oVoice, sGreeting, sIP, oNet, oFSO, oFile, sScript, errReturn
Dim intProcessId, oFolder

sIP = InputBox("Enter the machine name or IP:")
sGreeting = InputBox("Enter the voice message:")
Set oNet = CreateObject("WScript.Network")
oNet.MapNetworkDrive "X:", "\\" & sIP & "\C$"
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.CreateFolder ("X:\TempTemp")
Set oFile = oFSO.CreateTextFile ("X:\TempTemp\Voice.vbs", True)

sScript = "Set oVoice = WScript.CreateObject(" & Chr(34) & "SAPI.SpVoice" & Chr(34) & ")" & VbCrLf _
	& "oVoice.Speak " & Chr(34) & sGreeting & Chr(34)
oFile.Write sScript
oFile.Close
errReturn = GetObject("WinMgmts:\\" & sIP & "\root\cimv2:Win32_Process").Create("C:\windows\system32\wscript.exe C:\TempTemp\Voice.vbs", null, null, intProcessID)
WScript.Sleep 1000
oFSO.DeleteFolder "X:\TempTemp", True
oNet.RemoveNetworkDrive "X:", True