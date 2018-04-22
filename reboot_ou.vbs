'------------------------------------------------------------------------------
'
'  This script enumerates the computer accounts in a particular Active
'  Directory container and reboots all responsive systems.  The LDAP
'  path needs to be modified for the particular environment in which this 
'  script will run.  Also, the Win32Shutdown parameter can be set to a value
'  of 0 to force a logoff, 2 to "force" unresponsive applications to close
'  before rebooting, or 3 to power down the system.
'
'  Shawn Stugart
'=============================================================================


Const ADS_SCOPE_SUBTREE = 2
Set objShell = CreateObject("WScript.Shell")
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name, Location from 'LDAP://cn=computers,dc=nwtraders,dc=msft' " _
        & "where objectClass='computer'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Timeout") = 30 
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.Properties("Cache Results") = False 
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    strCompName = objRecordSet.Fields("name")
    Set objScriptExec = objShell.Exec("ping -n 2 -w 1000 " & strCompName)
    strPingStdOut = Lcase(objScriptExec.StdOut.ReadAll)
    If InStr(strPingStdOut, "reply from " & strComputer) Then
		Set objWMIService = GetObject("winmgmts:" & _
				"{impersonationLevel=impersonate,(Shutdown)}!\\" & _
				strCompName & "\root\cimv2")
		If Err.Number <> 0 Then
			WScript.Echo strCompName & ": " & Err.Description
			Err.Clear
		Else
			Set colOperatingSystems = objWMIService.ExecQuery _
					("SELECT * FROM Win32_OperatingSystem")
			For Each objOperatingSystem in colOperatingSystems
				objOperatingSystem.Win32Shutdown(1)
				Wscript.Echo "REBOOTING " & strCompName
			Next
		End If
	Else
		WScript.Echo strCompName & ": Host unreachable"
	End If
    objRecordSet.MoveNext
Loop