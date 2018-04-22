Option Explicit

Dim oRootDSE, oDomain, oCon
Dim oCmd, oRst, sUserName, sResult

Set oRootDSE = GetObject("LDAP://RootDSE")
Set oDomain = GetObject("LDAP://" _
  & oRootDSE.Get ("defaultNamingContext"))
Set oCon = CreateObject("ADODB.Connection") 
Set oCmd = CreateObject("ADODB.Command") 
oCon.Provider = "ADsDSOObject"  
oCon.Open 	
oCmd.ActiveConnection = oCon
sUserName = InputBox("Enter a portion of the name to search for:")
oCmd.CommandText = "<" & oDomain.ADsPath & ">;" _
	& "(&(objectCategory=person)(objectClass=user)(sn=*" & sUserName _
	& "*));cn,sAMAccountName;subTree"
	
Set oRst = oCmd.Execute()

On Error Resume Next

If oRst.EOF = True Then
	WScript.Echo "No records found."
Else
	Do While Not oRst.EOF
		sResult = sResult & oRst.Fields("cn") & ", " _
			& oRst.Fields("sAMAccountName") & vbCrLf
		oRst.MoveNext 
	Loop
	WScript.Echo "Records matching string " & Chr(34) _
	& sUserName & Chr(34) & ":" & vbCrLf & vbCrLf & sResult
End If