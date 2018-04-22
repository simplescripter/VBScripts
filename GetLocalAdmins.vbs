Option Explicit
Dim sComputer, arrAdminsToIgnore, oGroup, oUser, i
Dim sUser, sResult, sMembers, sReport, sFileName, iReply
Dim oFSO, oFile

CheckCScript

'On Error Resume Next

sFileName = "c:\clients2.txt"
arrAdminsToIgnore = Array("Admin3","Administrator")
SearchFile(sFileName) 'Uncomment EITHER SearchFile
'SearchAD				' OR SearchAD
WScript.Echo sReport
WriteIt

Sub SearchFile(sFile)
	Dim oFSO, oFile
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFSO.OpenTextFile(sFile,1)
	If Err.Number <> 0 Then
		WScript.Echo "Error opening " & sFile
		WScript.Quit
	End If
	Do Until oFile.AtEndOfStream
		Reporter(oFile.ReadLine)
	Loop
End Sub

Sub SearchAD()
	Dim oRootDSE, sDomain
	Dim oCon, oCom, oRS
	Set oRootDSE = GetObject("LDAP://RootDSE")
	sDomain = oRootDSE.Get("defaultNamingContext")
	Set oCon = CreateObject("ADODB.Connection") 
	Set oCom = CreateObject("ADODB.Command") 
	oCon.Provider = "ADsDSOObject"  
	oCon.Open 	
	oCom.ActiveConnection = oCon
	oCom.Properties("Page Size") = 1000
	oCom.CommandText =  "Select Name from 'LDAP://" & sDomain & "' " _
			& "Where objectClass='computer'"
	Set oRS = oCom.Execute() 
	Do While Not oRS.EOF
		Reporter(oRS.Fields("Name"))
		oRS.MoveNext
	Loop
End Sub

Function Reporter(sComputer)
	WScript.Echo "Trying " & sComputer & "..."
	If Pinger(sComputer) = 1 Then
		sMembers = sMembers & sComputer
		Set oGroup = GetObject("WinNT://" & sComputer & "/Administrators")
		If Err.Number = 0 Then
			For Each oUser In oGroup.Members
				sResult = 0
				sUser = oUser.Name
				For i = 0 To UBound(arrAdminsToIgnore)
					If sUser = arrAdminsToIgnore(i) Then
						sResult = sResult + 1
						Exit For
					Else
						'Do Nothing
					End If
				Next
				If sResult < 1 Then
					sMembers = sMembers & "," & sUser
				End If
			Next
			sReport = sReport & sMembers & VbCrLf
			sMembers = ""
		Else
			WScript.Echo "see this error?"
			Err.Clear
			sReport = sReport & sMembers & ",**ERROR**" & VbCrLf
		End If
	Else
		sReport = sReport & sComputer & "," & "**UNREACHABLE**" & VbCrLf
	End If
End Function

Function Pinger(sTarget)
	Dim iPingTimeOut, oShell, oScriptExec, sPingStdOut
	iPingTimeOut = "500"
	Set oShell = CreateObject("WScript.Shell")
	Set oScriptExec = oShell.Exec("ping -n 1 -w " & iPingTimeOut _
		 & " " & sTarget)
    sPingStdOut = Lcase(oScriptExec.StdOut.ReadAll)
    If InStr(sPingStdOut, "reply from") Then
		Pinger = 1
	Else
	    Pinger = 0
	End If
End Function

Sub CheckCScript()
	Dim oShell, hostname, iHelp
	Set oShell = CreateObject("WScript.Shell")
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    iHelp = MsgBox("At a command line, run " & Chr(34) _
	    	& "cscript " & WScript.ScriptFullName & Chr(34) _
	    	& vbCrLf & VbCrLf & "Do you want help?", vbYesNo)
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

Sub WriteIt()
	iReply = MsgBox("Do you want to write the results to file?", vbYesNo)
	If iReply = vbYes Then
		Dim sScriptPath, sReportPath
		sScriptPath = WScript.ScriptFullName
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		sReportPath = oFSO.GetParentFolderName(sScriptPath) & "\"
		Set oFile = oFSO.CreateTextFile(sReportPath & "AdminReport.csv", True)
		oFile.Write sReport
		oFile.Close
	Else
		WScript.Quit
	End If
End Sub