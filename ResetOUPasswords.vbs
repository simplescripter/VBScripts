Option Explicit

Dim oRootDSE, oDomain, oUser, oConnection, oCommand
Dim RS, sQuery, sOU, sNewPassword

sOU = "OU=YourOUHere,"
sNewPassword = "NewPasswordHere"

On Error Resume Next
Set oRootDSE = GetObject("LDAP://RootDSE")
oDomain = oRootDSE.Get("defaultNamingContext")
set oConnection = CreateObject("ADODB.Connection")
oConnection.Open "Provider=ADsDSOObject;"



sQuery = "<LDAP://" & sOU & oDomain & ">;(objectclass=user);adspath;subtree"

set oCommand = CreateObject("ADODB.Command")
oCommand.ActiveConnection = oConnection
oCommand.CommandText = sQuery

Set RS = oCommand.Execute
If RS.RecordCount = 0 Then
    WScript.Quit
Else
    Do Until RS.EOF
        Set oUser = GetObject(RS.Fields("adspath"))
	oUser.SetPassword sNewPassword
	oUser.setinfo
        RS.MoveNext
    Loop
    WScript.Echo "Password Reset Complete"
End If
