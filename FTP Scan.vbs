'---------------------------------------------------------------
'
'Shawn Stugart
'
'July 29th, 2005
'
'---------------------------------------------------------------





Option Explicit

const ForReading = 1
Dim strFileName, objFSO, objShell, objInputFile, ftpSite
Dim strServer, objIIS, obj, objScriptExec, strPingStdOut

On Error Resume Next

CheckForCScript

strFileName = "clients.txt"
set objFSO = CreateObject("Scripting.FileSystemObject")
set objShell = CreateObject("WScript.Shell")
set objInputFile = objFSO.OpenTextFile("c:\" & strFileName, ForReading)

Do Until objInputFile.AtEndOfStream
    strServer = objInputFile.Readline
    Set objScriptExec = objShell.Exec("ping -n 1 -w 250 " & strServer)
    strPingStdOut = Lcase(objScriptExec.StdOut.ReadAll)
    If InStr(strPingStdOut, "reply from ") Then
		set objIIS = GetObject("IIS://" & strServer & "/MSFTPSVC")
		If Err = 0 Then
			WScript.Echo "-----------------------------------------------------------------------------"
			Wscript.Echo "FTP Service FOUND on " & UCase(strServer)
			For each ftpSite in objIIS
				If (ftpSite.Class = "IIsFtpServer") then
					WScript.Echo vbTab & "FTP Site: " & ftpSite.ServerComment & " Anonymous Access Allowed = " _
						& ftpSite.AllowAnonymous
				End If
			next
			WScript.Echo "-----------------------------------------------------------------------------"
		Else
			WScript.Echo "FTP Service Not Found on " & strServer
			Err.Clear
		End If
	Else
	    WScript.Echo "***" & strServer & " is not available at this time.***"
	End If
Loop

Sub CheckForCScript
	dim hostname
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe" & vbCrLf & vbCrLf _
	    	& "At a command line, run " & chr(34) & "cscript " & WScript.ScriptFullName & chr(34)
	    WScript.Quit
	End If
End Sub
