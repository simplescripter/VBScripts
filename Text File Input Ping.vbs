'---------------------------------------------------------------------
'
'  Example script executing an external command (ping, in this case)
'  against a list of hosts found in a text file.
'
'  Shawn Stugart
'=====================================================================

const ForReading = 1
sFileName = InputBox("Enter File Name [clients.txt or servers.txt]", , "clients.txt")
set oFSO = CreateObject("Scripting.FileSystemObject")
set oShell = CreateObject("WScript.Shell")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)

Do Until oInputFile.AtEndOfStream
    strCompName = oInputFile.Readline
    Set objScriptExec = oShell.Exec("ping -n 1 -w 1000 " & strCompName)
    strPingStdOut = Lcase(objScriptExec.StdOut.ReadAll)
    If InStr(strPingStdOut, "reply from ") Then
		WScript.Echo strCompName & " is responding."
	Else
	    WScript.Echo strCompName & " is NOT responding."
	End If
Loop