'******************************************************************************
'
' The following script scans a list of computers and attempts to enumerate
' the user profiles on the system.  It looks for a file called LogonTime In
' the user's profile (created by the LogonTime.vbs logon script).
' The computer name, user's last logon time, profile name and profile size 
' is stored in a .csv report on the C: drive.  You'll need to
' be logged on with an Administrative account with privledges on the local
' machine (where the script is running) and on the machines being scanned.
'
' If you run this script with the /fast switch, it will not calculate the size
' of the user's profile, vastly speeding up the scan.
'
' Shawn Stugart
' 2-28-2006
'
' Updated on 4-11-2006
'******************************************************************************

Option Explicit

Const ForReading = 1

Dim hostname, iHelp, oShell, sReport, oReportFile
Dim sErrorBoxes, iErrors, sComplete
Dim oNet, sClients, oClientFile, sClient
Dim sProfilePath, oFSO, oFolder, dtmLogon
Dim oSubFolders, oSubFolder, sResult, oArgs, argFastScan
Dim sClientLogFile, oClientLogFile, iAnswer, sHomeDirPath

CheckCScript

iErrors = 0
sClients = "C:\Clients.txt" ' Location and name of your list of client machines
sReport = "C:\ProfileScanReport-" & FormatDateTime(Date,vbLongDate) & ".csv"
Set oNet = CreateObject("WScript.Network")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oArgs = WScript.Arguments.Named
argFastScan = LCase(oArgs.Item("fast"))
sProfilePath = "X:\Documents and Settings"
sHomeDirPath = "\\mlearning\HomeDir\"

On Error Resume Next

Set oClientFile = oFSO.OpenTextFile(sClients,1)
If Err.Number <> 0 Then
	WScript.Echo VbCrLf & VbCrLf & "***** Error opening " & sClients & ".  Make sure " _
		& "you've got a list of computers to scan in " & sClients & ". *****" & vbCrLf
	WScript.Quit
End If

Set oReportFile = oFSO.CreateTextFile(sReport, True)
If Err <> 0 Then
	WScript.Echo VbCrLf & VbCrLf & "***** Error creating report file " & sReport _
		& ".  Make sure the file is not currently in use and that you have permission " _
		& "to create the report in the specified directory. *****" & vbCrLf
	WScript.Quit
End If

oReportFile.WriteLine("Computer,Profile,LastLogon,Size(in GB)")

Do Until oClientFile.AtEndOfStream
	sClient = oClientFile.ReadLine
	oNet.RemoveNetworkDrive "X:", True
	If Err.Number <> 0 Then
		Err.Clear
	End If
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
		    		sClientLogFile = sProfilePath & "\" & oSubFolder.Name & "\LogonTime"
		    		Set oClientLogFile = oFSO.OpenTextFile(sClientLogFile,ForReading)
		    		If Err.Number = 0 Then
			    		dtmLogon = oClientLogFile.ReadLine
			    		oClientLogFile.Close
			    	Else
			    		Err.Clear
					sClientLogFile = sHomeDirPath & oSubFolder.Name & "\LogonTime"
		    			Set oClientLogFile = oFSO.OpenTextFile(sClientLogFile,ForReading)
					If Err.Number = 0 Then
						dtmLogon = oClientLogFile.ReadLine
			    			oClientLogFile.Close
					Else
						dtmLogon = "NO_HISTORY"
			    			Err.Clear
					End If
			    	End If
				If WScript.Arguments.Named.Exists("fast") Then
				    sResult = sResult & dtmLogon & ", FAST_SCAN"
				Else
		    		    sResult = sResult & dtmLogon & "," & Round((oSubFolder.Size / 1073741824),3)
				End If
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
	sComplete = sComplete & "Could not scan the following " & iErrors & " machine(s):" & VbCrLf & VbCrLf _
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