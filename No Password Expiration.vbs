' *****************************************************************************
' Disables user account password expiration on any user account starting with
' "User" or "Test".  Written for the Network+ class. 
' 
'
' Author: Shawn Stugart
' Date: 10-11-2005
'******************************************************************************

Option Explicit

Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
 
Dim oRootDSE, oDomain, oCon, oCmd
Dim oRst, sUser, oUser

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
	& "(&(objectCategory=person)(objectClass=user)(|(name=User*)(name=Test*)));sAMAccountName,distinguishedName;" _
	& "subTree"
	
Set oRst = oCmd.Execute()

Do While Not oRst.EOF
	SetPassProp(oRst.Fields("distinguishedName"))
	oRst.MoveNext 
Loop

MsgBox "Operation Complete."

Sub SetPassProp(sUser)
	Dim oUser, iUAC 
	Set oUser = GetObject("LDAP://" & sUser)
	iUAC = oUser.Get("userAccountControl")
	If ADS_UF_DONT_EXPIRE_PASSWD AND iUAC Then
    	Wscript.Echo "Already enabled on " & oRst.Fields("sAMAccountName")
	Else
	    oUser.Put "userAccountControl", iUAC XOR _
	        ADS_UF_DONT_EXPIRE_PASSWD
	    oUser.SetInfo
	    WScript.Echo "Password never expires is now enabled for " & oRst.Fields("sAMAccountName")
	End If
End Sub

Sub CheckCScript()
	Dim hostname, iHelp, oShell
	Set oShell = CreateObject("WScript.Shell")
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