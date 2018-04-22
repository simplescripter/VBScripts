Option Explicit

Dim oConnection, oRootDSE, oDomainADSI, oCommand
Dim oRecordSet, sADsPath, oFSO, oFile, sFile, sName

sFile = "C:\EnabledUserReport.csv"
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.CreateTextFile(sFile, True)
oFile.WriteLine "First,Middle,Last"

Set oConnection = CreateObject("ADODB.Connection")
oConnection.Open "Provider=ADsDSOObject;"
Set oRootDSE = GetObject("LDAP://RootDSE")
Set oDomainADSI = GetObject("LDAP://" _
  & oRootDSE.Get ("defaultNamingContext"))
Set oCommand = CreateObject("ADODB.Command")
oCommand.ActiveConnection = oConnection
oCommand.Properties("Page Size") = 100
oCommand.CommandText = "<" & oDomainADSI.ADsPath & ">;" _
       & "(&(objectCategory=Person)(objectClass=user)(!userAccountControl=514)(sN=*));" _
           & "givenName,initials,sN;subtree"
 
Set oRecordSet = oCommand.Execute
oRecordSet.Movefirst
Do Until oRecordSet.EOF
    sName = oRecordSet.Fields("givenName").Value & ","
    sName = sName & oRecordSet.Fields("initials").Value & ","
    sName = sName & oRecordSet.Fields("sN").Value
    oFile.WriteLine sName
    sName = ""
    oRecordSet.MoveNext
Loop

oConnection.Close
