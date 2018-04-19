'Example Script that pings every machine listed in an AD container
'
'Written by Shawn Stugart
'========================================================================
Const ADS_SCOPE_SUBTREE = 2
set objShell = CreateObject("WScript.Shell")
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name from 'LDAP://CN=Computers,DC=domain100,DC=internal' " _
        & "where objectClass='computer'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Timeout") = 30 
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.Properties("Cache Results") = False 
Set objRecordSet = objCommand.Execute

Do Until objRecordSet.EOF
    strCompName = objRecordSet.Fields("Name")
    Set objScriptExec = objShell.Exec("ping -n 1 -w 1000 " & strCompName)
    strPingStdOut = Lcase(objScriptExec.StdOut.ReadAll)
    If InStr(strPingStdOut, "reply from ") Then
		WScript.Echo strCompName & " is responding."
	Else
	    WScript.Echo strCompName & " is NOT responding."
	End If
    objRecordSet.MoveNext
Loop