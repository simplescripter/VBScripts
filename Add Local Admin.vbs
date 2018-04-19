' ****************************************************************
' This script binds to the local SAM database of each system 
' listed in C:\clients.txt and creates a local administrator
' account called Admin with a password of P@ssword
' 
' Shawn Stugart
' 
' 8-22-2005
' ****************************************************************

Option Explicit

const ForReading = 1
dim sFileName, oFSO, oInputFile, oUser, oDom
dim sComputer, oScriptExec, oShell, oAdmin, oGroup
Dim sPingStdOut, sNewPassword, hostname

On Error Resume Next

Call CheckCScript
sNewPassword = "P@ssw0rd"
sFileName = "C:\clients.txt"
Set oShell = CreateObject("WScript.Shell")
set oFSO = CreateObject("Scripting.FileSystemObject")
set oInputFile = oFSO.OpenTextFile(sFileName, ForReading)
Do Until oInputFile.AtEndOfStream
	sComputer = oInputFile.Readline
	Set oScriptExec = oShell.Exec("ping -n 1 -w 1000 " & sComputer)
    sPingStdOut = Lcase(oScriptExec.StdOut.ReadAll)
    If InStr(sPingStdOut, "reply from ") Then
		Set oDom = GetObject("WinNT://" & sComputer)
		If Err.Number = 0 Then
			Set oUser = oDom.Create("user","Admin")
			oUser.SetInfo
			Set oGroup = GetObject("WinNT://" & sComputer & "/Administrators,group")
			oGroup.Add(oUser.ADsPath)
		    oUser.SetPassword sNewPassword
		    oUser.AccountDisabled = False
		    oUser.SetInfo
		    WScript.StdOut.WriteLine "Admin account created and added to Administrators group on " & sComputer
		Else
		    WScript.StdOut.WriteLine "FAILED to create Admin account on " & sComputer
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