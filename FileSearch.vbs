CheckCScript

sFileName = "X:\PowerShell\test.txt"
sComputerList = "C:\clients.txt"
sReportName = "C:\FileScan.txt"

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.OpenTextFile(sComputerList, 1)
Set oNet = CreateObject("WScript.Network")
On Error Resume Next
Do Until oFile.AtEndOfStream
	sComp = oFile.ReadLine
	oNet.RemoveNetworkDrive "X:"
	If Err.Number <> 0 Then
		Err.Clear
	End If
	oNet.MapNetworkDrive "X:", "\\" & sComp & "\C$"
	If Err.Number <> 0 Then
		sResult = sResult & "ERROR CONNECTING TO " & sComp & vbCrLf
		WScript.Echo "ERROR CONNECTING TO " & sComp
		Err.Clear
	Else
		If oFSO.FileExists(sFileName) Then
			sResult = sResult & sComp & ": Found file" & VbCrLf
			WScript.Echo sComp & ": Found file"
		Else
			sResult = sResult & "File doesn't exist on " & sComp & VbCrLf
			WScript.Echo "File doesn't exist on " & sComp
		End If
	End If
Loop

Set oReport = oFSO.CreateTextFile(sReportName, true)
oReport.Write Now & vbCrLf & vbCrLf & sResult
WScript.Echo "----------------------------------------------------------------"
WScript.Echo "Done. The results of this scan have been saved to " & sReportName

Sub CheckCScript()
	Dim oShell, hostname, iHelp
	Set oShell = CreateObject("WScript.Shell")
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    iHelp = MsgBox("At a command line, run " & Chr(34) _
	    	& "cscript " & WScript.ScriptFullName & Chr(34) _
	    	& vbCrLf & "Do you want help?", vbYesNo)
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