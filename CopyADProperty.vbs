Option Explicit

Dim oRootDSE, oDomain, oCon, oCmd, oRst, sDN
Dim oUser, sPropCopyFrom, sPropCopyTo, sProp

On Error Resume Next

sPropCopyFrom = "userPrincipalName"
sPropCopyTo = "company"

Set oRootDSE = GetObject("LDAP://RootDSE")
Set oDomain = GetObject( "LDAP://" & oRootDSE.Get("defaultNamingContext"))
Set oCon = CreateObject("ADODB.Connection") 
Set oCmd = CreateObject("ADODB.Command") 
oCon.Provider = "ADsDSOObject"  
oCon.Open 	
oCmd.ActiveConnection = oCon
oCmd.CommandText = "<" & oDomain.ADsPath & ">;" _
	& "(&(objectCategory=Person)(objectClass=user));distinguishedName;subTree"
oCmd.Properties("Page Size") = 100
oCmd.Properties("Timeout") = 60
oCmd.Properties("Cache Results") = False
Set oRst = oCmd.Execute()
Do While Not oRst.EOF
	sDN = oRst.Fields("distinguishedName")
	Set oUser = GetObject("LDAP://" & sDN)
	sProp = oUser.Get(sPropCopyFrom)
	sProp = Split(sProp, "@")
 	oUser.Put sPropCopyTo, sProp(0)
 	If Err.Number <> 0 Then
 		Err.Clear
 	End If
 	oUser.SetInfo
 	oRst.MoveNext
Loop
WScript.Echo "Done."
 
	