Sub CheckCScript()
	Dim oShell, hostname, iHelp
	Set oShell = CreateObject("WScript.Shell")
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    iHelp = MsgBox("At a command line, run " & Chr(34) _
	    	& "cscript " & WScript.ScriptFullName & Chr(34) _
	    	& vbCrLf & vbCrLf & "Optionally, you can include the /fast " _
		& "switch, in which case the size of the user's profile " _
		& "will not be calculated, significantly speeding up " _
		& "the scan." & vbCrLf & vbCrLf & "Do you want help?", vbYesNo)
	    If iHelp = vbYes Then
	    	oShell.Run("cmd.exe")
	    	WScript.Sleep 500
	    	oShell.SendKeys "cscript.exe " & Chr(34) & WScript.ScriptFullName & Chr(34)
	    	WScript.Quit
	    Else
	    	WScript.Quit
	    End If
	End If
End Sub