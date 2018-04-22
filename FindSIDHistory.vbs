Option Explicit

Dim oRootDSE, oDomain, oCon
Dim oCmd, oRst

CheckCScript

Set oRootDSE = GetObject("LDAP://RootDSE")
Set oDomain = GetObject("LDAP://" _
  & oRootDSE.Get ("defaultNamingContext"))
Set oCon = CreateObject("ADODB.Connection") 
Set oCmd = CreateObject("ADODB.Command") 
oCon.Provider = "ADsDSOObject"  
oCon.Open 	
oCmd.ActiveConnection = oCon
oCmd.CommandText = "<" & oDomain.ADsPath & ">;" _
	& "(&(objectCategory=person)(objectClass=user)(sIDHistory=*));sAMAccountName, distinguishedName;" _
	& "subTree"
	
Set oRst = oCmd.Execute()

On Error Resume Next

WScript.Echo "The following accounts have SIDHistory set:" & vbCrLf

Do While Not oRst.EOF
	WScript.Echo vbTab & FetchSIDHistory(oRst.Fields("sAMAccountName"), _
		oRst.Fields("distinguishedName"))
	oRst.MoveNext 
Loop

WScript.Echo VbCrLf & VbCrLf & "Finished scanning accounts."

Sub CheckCScript()
	Dim hostname, iHelp, oShell
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

Function FetchSIDHistory(sUser,sDN)
	Dim oUser, vArray, vSID, sSIDHistory
	Set oUser = GetObject("LDAP://" & sDN) 
    vArray = oUser.GetEx("sIDHistory")
    For Each vSID in vArray
    	vSID = OctetToHexStr(vSID)
    	sSIDHistory = sSIDHistory & vSID & VbCrLf
    Next
    FetchSIDHistory = sUser & vbTab & sSIDHistory
End Function

Function OctetToHexStr(sOctet)
  	Dim k
  	OctetToHexStr = ""
 	For k = 1 To Lenb(sOctet)
    	OctetToHexStr = OctetToHexStr _
      	& Right("0" & Hex(Ascb(Midb(sOctet, k, 1))), 2)
 	Next
End Function