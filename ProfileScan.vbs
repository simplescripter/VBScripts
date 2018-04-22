'******************************************************************************
'
' The following script scans a list of computers and attempts to enumerate
' the user profiles on the system with a Last Access date.
' The information is stored in a .csv report on the C: drive.  You'll need To
' be logged on with an Administrative account with priveledges on the local
' machine (where the script is running) and on the machines being scanned.
'
' Shawn Stugart
' 2-9-2006
'
'******************************************************************************

Option Explicit

Dim hostname, iHelp, oShell, sReport, oReportFile
Dim sErrorBoxes, iErrors, sComplete
Dim oNet, sClients, oClientFile, sClient
Dim sProfilePath, oFSO, oFolder
Dim oSubFolders, oSubFolder, sResult
Dim oNTUserFile, dtmAccessDate, dtmModDate, iAnswer

CheckCScript

iErrors = 0
sClients = "C:\Clients.txt" ' Location and name of your list of client machines
sReport = "C:\ProfileScanReport-" & FormatDateTime(Date,vbLongDate) & ".csv"
Set oNet = CreateObject("WScript.Network")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oClientFile = oFSO.OpenTextFile(sClients,1)
Set oReportFile = oFSO.CreateTextFile(sReport, True)
sProfilePath = "X:\Documents and Settings"

On Error Resume Next

oReportFile.WriteLine("Computer,Profile,DateLastAccessed,DateLastModified,Size(in GB)")

Do Until oClientFile.AtEndOfStream
	sClient = oClientFile.ReadLine
	oNet.MapNetworkDrive "X:", "\\" & sClient & "\C$"
	If Err.Number = 0 Then
		WScript.StdOut.Write(".")
		Set oFolder = oFSO.GetFolder(sProfilePath)
		Set oSubFolders = oFolder.SubFolders
		For Each oSubFolder in oSubFolders
		    Select Case oSubFolder.Name
			Case "All Users"
			Case "Default User"
			Case "LocalService"
			Case "NetworkService"
		    	Case Else
		    		sResult = sResult & sClient & ","
		    		sResult = sResult & oSubFolder.Name & ","
		    		Set oNTUserFile = oFSO.GetFile(oSubFolder & "\NTUSER.DAT")
		    		dtmAccessDate = oNTUserFile.DateLastAccessed
		    		dtmModDate = oNTUserFile.DateLastModified
		    		sResult = sResult & dtmAccessDate & "," & dtmModDate & "," & oSubFolder.Size / 1073741824
		            oReportFile.WriteLine sResult
		            sResult = ""
		    End Select
		Next
	Else
		WScript.StdOut.Write(".")
		iErrors = iErrors + 1
		sErrorBoxes = sErrorBoxes & sClient & VbCrLf
		Err.Clear
	End If
	oNet.RemoveNetworkDrive "X:", True
Loop

If iErrors > 0 Then
	sComplete = sComplete & "Could not scan the following " & iErrors & " machines:" & VbCrLf & VbCrLf _
		& sErrorBoxes & VbCrLf & VbCrLf & "The following report has been generated: " & sReport _
		& VbCrLf & VbCrLf & "Would you like to open the report now?"
Else
	sComplete = sComplete & "Scanned all systems successfully." & VbCrLf & VbCrLf _
		& "The following report has been generated: " & sReport _
		& VbCrLf & VbCrLf & "Would you like to open the report now?"
End If

iAnswer = MsgBox(sComplete,vbYesNo)
	If iAnswer = vbYes Then
		oShell.Run "excel.exe " & Chr(34) & sReport & Chr(34)
	Else
		WScript.Quit
	End If

Sub CheckCScript()
	Set oShell = CreateObject("WScript.Shell")
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    iHelp = MsgBox("At a command line, run " & Chr(34) _
	    	& "cscript " & WScript.ScriptFullName & Chr(34) _
	    	& vbCrLf & vbCrLf & "Do you want help?", vbYesNo)
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