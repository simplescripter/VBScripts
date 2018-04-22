'---------------------------------------------------------------
'	This script uses a list of computers found in Active Directory
'   to scan for the Microsoft FTP Service.  If the service is
'   found on a system, the script reports on the AllowAnonymous
'   value of each FTP site hosted by the system.
'
'Shawn Stugart
'
'July 29th, 2005
'
'---------------------------------------------------------------


Option Explicit

Dim objFSO, objShell, objInputFile, ftpSite
Dim strServer, objIIS, obj, objScriptExec, strPingStdOut
Dim objConnection, objCommand, objRecordSet

On Error Resume Next

CheckForCScript

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name from 'LDAP://CN=Computers,DC=nwtraders,DC=msft' " _
        & "where objectClass='computer'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Timeout") = 30 
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.Properties("Cache Results") = False 
Set objRecordSet = objCommand.Execute
set objShell = CreateObject("WScript.Shell")

Do Until objRecordSet.EOF
    strServer = objRecordSet.Fields("Name")
    Set objScriptExec = objShell.Exec("ping -n 1 -w 250 " & strServer)
    strPingStdOut = Lcase(objScriptExec.StdOut.ReadAll)
    If InStr(strPingStdOut, "reply from ") Then
		set objIIS = GetObject("IIS://" & strServer & "/MSFTPSVC")
		If Err = 0 Then
			WScript.Echo "-----------------------------------------------------------------------------"
			Wscript.Echo "FTP Service FOUND on " & UCase(strServer)
			For each ftpSite in objIIS
				If (ftpSite.Class = "IIsFtpServer") then
					WScript.Echo vbTab & "FTP Site: " & ftpSite.ServerComment & " Anonymous Access Allowed = " _
						& ftpSite.AllowAnonymous
				End If
			next
			WScript.Echo "-----------------------------------------------------------------------------"
		Else
			WScript.Echo "FTP Service Not Found on " & strServer
			Err.Clear
		End If
	Else
	    WScript.Echo "***" & strServer & " is not available at this time.***"
	End If
	objRecordSet.MoveNext
Loop

Sub CheckForCScript
	dim hostname
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe" & vbCrLf & vbCrLf _
	    	& "At a command line, run " & chr(34) & "cscript " & WScript.ScriptFullName & chr(34)
	    WScript.Quit
	End If
End Sub
