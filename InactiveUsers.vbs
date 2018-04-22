'******************************************************************************
' Adapted from a script by Don Jones and another from RLMueller.net
'
' This script will currently report on and optionally disable user accounts
' that have exceeded a certain period of inactivity (in weeks) as defined by 
' the iNumberOfWeeks variable.  Run at the command line with cscript.exe, And
' use /r to report only, /d to disable inactive or never used accounts, /r /d 
' to both report and disable.
'
' The current version of this script does not support the /m switch (moving
' inactive users to their own OU)
'
' Shawn Stugart, 9-26-2006
'
'******************************************************************************

Dim dDate, oUser, oObject, oGroup, oArgs
Dim iFlags, iDiff, iResult, iNumberOfWeeks
Dim oRootDSE, oDomain, oCon, oCmd, oRst
Dim oShell, sLog, lngBiasKey
Dim strDN, lngDate, objDate, dtmDate, lngHigh, lngLow
Const UF_ACCOUNTDISABLE = &H0002

'The following variables need to be set properly for the script to work
'******************************************************************************
iNumberOfWeeks = 1 					' Set maximum number of weeks for user inactivity
sOU = ""								' Enter LDAP path to the OU or container to scan
								 		' If you leave the path empty, then entire domain
								 		' will be scanned
sDisabledUsersOU = ""
'******************************************************************************
CheckCScript

Set oArgs = WScript.Arguments.Named
If oArgs.Count < 1 Then
	Help
End If

' Obtain local Time Zone bias from machine registry. Thanks to RLMueller.net!
Set oShell = CreateObject("Wscript.Shell")
lngBiasKey = oShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
  	& "TimeZoneInformation\ActiveTimeBias")
If UCase(TypeName(lngBiasKey)) = "LONG" Then
  	lngBias = lngBiasKey
ElseIf UCase(TypeName(lngBiasKey)) = "VARIANT()" Then
  	lngBias = 0
  	For k = 0 To UBound(lngBiasKey)
    	lngBias = lngBias + (lngBiasKey(k) * 256^k)
  	Next
End If

If oArgs.Exists("r") Then
	sLog = "Report generated on " & Now & VbCrLf & VbCrLf & "The following " _
		& "user accounts have been inactive for at least " & iNumberOfWeeks _
		& " weeks:" & VbCrLf
End If
Set oShell = CreateObject("WScript.Shell")
If sOU = "" Then
	Set oRootDSE = GetObject("LDAP://RootDSE")
	Set oDomain = GetObject("LDAP://" _
	  & oRootDSE.Get ("defaultNamingContext"))
	Set oCon = CreateObject("ADODB.Connection") 
	Set oCmd = CreateObject("ADODB.Command") 
	oCon.Provider = "ADsDSOObject"  
	oCon.Open 	
	oCmd.ActiveConnection = oCon
	oCmd.CommandText = "<" & oDomain.ADsPath & ">;" _
		& "(&(objectCategory=Person)(objectClass=user));distinguishedName," _
		& "lastLogon;subTree"
	oCmd.Properties("Page Size") = 100
	oCmd.Properties("Timeout") = 60
	oCmd.Properties("Cache Results") = False
	Set oRst = oCmd.Execute()
	Do While Not oRst.EOF
		strDN = oRst.Fields("distinguishedName")
      	lngDate = oRst.Fields("lastLogon")
      	On Error Resume Next
      	Set objDate = lngDate 'Thanks again to RLMueller.net for the time conversion!
      	If Err.Number <> 0 Then
        	On Error GoTo 0
        	dtmDate = #1/1/1601#
      	Else
        	On Error GoTo 0
        	lngHigh = objDate.HighPart
        	lngLow = objDate.LowPart
        	If lngLow < 0 Then
          		lngHigh = lngHigh + 1
        	End If
        	If (lngHigh = 0) And (lngLow = 0 ) Then
          		dtmDate = #1/1/1601#
       		Else
          		dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
            		+ lngLow)/600000000 - lngBias)/1440
        	End If
      	End If
		GetUser oRst.Fields("distinguishedName"), dtmDate
		oRst.MoveNext 
	Loop
Else
	Set oCon = CreateObject("ADODB.Connection") 
	Set oCmd = CreateObject("ADODB.Command") 
	oCon.Provider = "ADsDSOObject"  
	oCon.Open 	
	oCmd.ActiveConnection = oCon 	
	oCmd.CommandText = "<LDAP://" & sOU & ">;" _
		& "(&(objectCategory=Person)(objectClass=user));distinguishedName," _
		& "lastLogon;subTree"
	oCmd.Properties("Page Size") = 100
	oCmd.Properties("Timeout") = 60
	oCmd.Properties("Cache Results") = False
	Set oRst = oCmd.Execute()
	Do While Not oRst.EOF
		strDN = oRst.Fields("distinguishedName")
      	lngDate = oRst.Fields("lastLogon")
      	On Error Resume Next
      	Set objDate = lngDate
      	If Err.Number <> 0 Then
        	On Error GoTo 0
        	dtmDate = #1/1/1601#
      	Else
        	On Error GoTo 0
        	lngHigh = objDate.HighPart
        	lngLow = objDate.LowPart
        	If lngLow < 0 Then
          		lngHigh = lngHigh + 1
        	End If
        	If (lngHigh = 0) And (lngLow = 0 ) Then
          		dtmDate = #1/1/1601#
        	Else
          		dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
            	+ lngLow)/600000000 - lngBias)/1440
       		End If
      	End If
		GetUser oRst.Fields("distinguishedName"),dtmDate
		oRst.MoveNext 
	Loop
End If
	
On error resume Next

Sub CheckCScript()
	Dim hostname, iHelp
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    iHelp = MsgBox("At a command line, run " & Chr(34) _
	    	& "cscript " & WScript.ScriptFullName & Chr(34) _
	    	& vbCrLf & vbCrLf & "You must run the script with one " _
			& "or more of the following switches: " & VbCrLf & vbTab _
			& "/r" & vbTab & "Generates a report in the local Application log" & VbCrLf _
			& vbTab & "/d" & vbTab & "Disables user accounts that have been inactive " _
			& "for more than " & iNumberOfWeeks & " weeks" & VbCrLf _
			& vbTab & "/m" & vbTab & "Moves inactive users to the " & sDisabledUsersOU _
			& " container", vbOK)
		WScript.Quit
	End If
End Sub

Sub Help()
	WScript.Echo "You must run this script with one or more of the following " _
		& "switches: " & VbCrLf & vbTab & "/r" & vbTab & "Generates a report " _
		& "in the local Application log" & VbCrLf & vbTab & "/d" & vbTab _
		& "Disables user accounts that have been inactive for more than " _
		& iNumberOfWeeks & " weeks" & VbCrLf & vbTab & "/m" & vbTab & "Moves " _
		& "inactive users to the " & sDisabledUsersOU & " container"
		WScript.Quit
End Sub
Function GetUser(sUser, dLastLogon)
	On Error Resume Next
    Set oUser = GetObject("LDAP://" & sUser)
    If dLastLogon = "1/1/1601" Then
    	If oArgs.Exists("d") Then 
	    	If DisableUser(sUser) = 1 Then
	     		sLog = sLog & oUser.Name & " has been DISABLED" & VbCrLf
	     	Else
	     		sLog = sLog & "An error occurred disabling " & oUser.Name
	     	End If
	    Else
	    	sLog = sLog & oUser.Name & " has NEVER logged on" & VbCrLf
    	End If
    Else
    	dLastLogon = CDate(dLastLogon)
    	dLastLogon = FormatDateTime(dLastLogon,vbShortDate)
    	'calculate how long ago that was in weeks
    	iDiff = DateDiff("ww", dLastLogon, Now)
	    If iDiff >= iNumberOfWeeks Then
	    	If oArgs.Exists("d") Then 
	    		If DisableUser(sUser) = 1 Then
	     			sLog = sLog & oUser.Name & " has been DISABLED" & VbCrLf
	     		Else
	     			sLog = sLog & "An error occurred disabling " & oUser.Name
	     		End If
	     	Else
	     		sLog = sLog & oUser.Name & " has been inactive since " _
	     			& dLastLogon & VbCrLf
	     	End If
	  	End If
	 End If
End Function

Function DisableUser(sUser)
	'Set oUser = GetObject("LDAP://" & sUser)
	If oUser.AccountDisabled = False Then
		oUser.AccountDisabled = True
		oUser.SetInfo
	End If
	If Err.Number = 0 Then
		DisableUser = 1
	Else
		Err.Clear
		DisableUser = 0
	End If
End Function

If oArgs.Exists("r") Then
	oShell.LogEvent 0, sLog
End If