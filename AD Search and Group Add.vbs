Option Explicit
Dim oConnection, oCommand, oResult, oGroup, sGroup, sSearchBase, sProperty

sSearchBase = "DC=Fourthcoffee,DC=com"
sGroup = "CN=Test,OU=Marketing,DC=Fourthcoffee,DC=com"
sProperty = "telephoneNumber"

On Error Resume Next
Set oConnection = CreateObject("ADODB.Connection")
Set oCommand = CreateObject("ADODB.Command")
oConnection.Provider = "ADsDSOObject"
oConnection.Open
oCommand.ActiveConnection = oConnection
oCommand.CommandText = "<LDAP://" & sSearchBase & ">;" _
	& "(&(objectClass=user)(objectCategory=person)(" & sProperty & "=*));" _
	& "ADsPath;" _
	& "subTree"
Set oResult = oCommand.Execute()
Set oGroup = GetObject("LDAP://" & sGroup)
Do While Not oResult.EOF
	oGroup.Add(oResult.Fields("ADsPath"))
	oResult.MoveNext
Loop

