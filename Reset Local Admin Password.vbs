Option Explicit

const ForReading = 1
dim sAdminName, sFileName, oFSO, oInputFile
dim sComputer, oScriptExec, oShell, oAdmin
Dim sPingStdOut, sNewPassword, hostname

On Error Resume Next

Call CheckCScript
sNewPassword = "BigF00+"
sAdminName = "StudentX"
sFileName = "C:\clients.txt"
Set oShell = CreateObject("WScript.Shell")
set oFSO = CreateObject("Scripting.FileSystemObject")
set oInputFile = oFSO.OpenTextFile(sFileName, ForReading)
Do Until oInputFile.AtEndOfStream
	sComputer = oInputFile.Readline
	Set oScriptExec = oShell.Exec("ping -n 1 -w 1000 " & sComputer)
    sPingStdOut = Lcase(oScriptExec.StdOut.ReadAll)
    If InStr(sPingStdOut, "reply from ") Then
		Set oAdmin = GetObject("WinNT://" & sComputer & "/" & sAdminName)
		If Err.Number = 0 Then
		    oAdmin.SetPassword sNewPassword
		    oAdmin.SetInfo
		    WScript.StdOut.WriteLine sAdminName & " password set on " & sComputer
		Else
		    WScript.StdOut.WriteLine "FAILED to set " & sAdminName & " password on " & sComputer
		    Err.Clear
		End If 
    Else
        WScript.StdOut.WriteLine "...............NO RESPONSE FROM " & sComputer
    End If
Loop

'******************************************************************************************************
Sub CheckCScript
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe." & vbCrLf & vbCrLf _
	        & "At a command prompt, use " & """cscript " & WScript.ScriptName & """"
	    WScript.Quit
	End If
End Sub