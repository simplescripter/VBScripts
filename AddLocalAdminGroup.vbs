Option Explicit

Dim sDomainGlobalGroup, oRootDSE, oShell, sOU
Dim oDomInfo, sDomainName, oCon, oCmd, oRst
Dim oScriptExec, sPingStdOut, sComputer
Dim oDomainGroup, oLocalAdminGroup, hostname

sDomainGlobalGroup = "Admins3"

'On Error Resume Next

CheckCscript

Set oRootDSE = GetObject("LDAP://RootDSE")
Set oShell = CreateObject("WScript.Shell")
' Replace CN=Computers in the following line with the OU containing
' the servers on which you want to add the new Administrator group

sOU = "<LDAP://CN=Computers," & oRootDSE.Get("defaultNamingContext") & ">"
Set oDomInfo = CreateObject("ADSystemInfo")
sDomainName = oDomInfo.DomainShortName
Set oCon = CreateObject("ADODB.Connection") 
Set oCmd = CreateObject("ADODB.Command") 
oCon.Provider = "ADsDSOObject"  
oCon.Open 	
oCmd.ActiveConnection = oCon
oCmd.CommandText = sOU & ";(objectCategory=Computer);name;subTree"
	
Set oRst = oCmd.Execute()

Do While Not oRst.EOF
	AddAdmin(oRst.Fields("name"))
	oRst.MoveNext 
Loop

Sub AddAdmin(sComputer)
    Set oScriptExec = oShell.Exec("ping -n 1 -w 500 " & sComputer)
    sPingStdOut = Lcase(oScriptExec.StdOut.ReadAll)
    If InStr(sPingStdOut, "reply from ") Then
		Set oLocalAdminGroup = GetObject("WinNT://" _
			& sComputer & "/Administrators,group")
		Set oDomainGroup = GetObject("WinNT://" & sDomainName & "/" & sDomainGlobalGroup)
		oLocalAdminGroup.Add(oDomainGroup.ADsPath)
		If Err.Number = 0 Then
			WScript.Echo sComputer & " done."
		Else
			WScript.Echo "FAILED ADDING " & sDomainGlobalGroup & " GROUP TO " & sComputer
			Err.Clear
		End If
	Else
		WScript.Echo "Could not contact " & sComputer & "."
	End If
End Sub

Sub CheckCScript()
	hostname = lcase(right(WSCript.Fullname, 11))
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
