Option Explicit

Dim oFSO, oRootDSE, oDomain, oCon
Dim oCmd, oRst, iDone, oShell, oFile
Dim arrMail, sMail, sMailAddresses

CheckCScript


Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
Set oFile = oFSO.CreateTextFile("C:\emails.txt",True)
Set oRootDSE = GetObject("LDAP://RootDSE")
Set oDomain = GetObject("LDAP://" _
  & oRootDSE.Get ("defaultNamingContext"))
Set oCon = CreateObject("ADODB.Connection") 
Set oCmd = CreateObject("ADODB.Command") 
oCon.Provider = "ADsDSOObject"  
oCon.Open 	
oCmd.ActiveConnection = oCon
oCmd.CommandText = "<" & oDomain.ADsPath & ">;" _
	& "(&(objectCategory=person)(objectClass=user)(proxyAddresses=*));proxyAddresses;" _
	& "subTree"
	
Set oRst = oCmd.Execute()

'On Error Resume Next

Do While Not oRst.EOF
	arrMail = oRst.Fields("proxyAddresses").Value
	For Each sMail In arrMail
		sMailAddresses = sMailAddresses & sMail & ","
	Next
	oFile.WriteLine(sMailAddresses)
	sMailAddresses = Null
	oRst.MoveNext 
Loop

oFile.Close
iDone = MsgBox("Done. The results have been saved as C:\emails.txt." _
	& " Do you want to view this file?",vbYesNo)
If iDone = vbYes Then
	oShell.Run "notepad C:\emails.txt"
End If

Sub CheckCScript()
	Dim hostname, iHelp
	hostname = lcase(Right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    iHelp = MsgBox("At a command line, run " & Chr(34) _
	    	& "cscript " & WScript.ScriptFullName & Chr(34) _
	    	& vbCrLf & vbCrLf & "Do you want help?", vbYesNo)
	    If iHelp = vbYes Then
	    	Set oShell = CreateObject("WScript.Shell")
	    	oShell.Run("cmd.exe")
	    	WScript.Sleep 500
	    	oShell.SendKeys "cscript.exe " & Chr(34) & WScript.ScriptFullName & Chr(34)
	    	WScript.Quit
	    Else
	    	WScript.Quit
	    End If
	End If
End Sub