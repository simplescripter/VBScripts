Option Explicit

Const SEC_IN_DAY = 86400
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000

Dim iExpiryNotify, sDomain, oDomain, iMaxPwdAge
Dim oUser, iCurrentValue, iNoExpire, iExpiring
Dim sResult, dtmValue, iTimeInterval, iExpired, sLog
Dim sExpires, oRootDSE, oDomainADSI, oCon, oCmd, oShell, sReport
Dim oRst, oFSO, oReport, iDays, sUser, iUserNo, iNotExpiring

iExpiryNotify = int(41) 'number of days before password expiration to notify
sDomain = "classnet" 'set domain name here

iNoExpire = 0
iExpiring = 0
iExpired = 0
iNotExpiring = 0

sReport = "C:\PasswordReport-" & FormatDateTime(Date,vbLongDate) & ".csv"
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oReport = oFSO.CreateTextFile(sReport,True)
Set oDomain = GetObject("WinNT://" & sDomain)
iMaxPwdAge = oDomain.Get("MaxPasswordAge")
iMaxPwdAge = (iMaxPwdAge/SEC_IN_DAY)

Set oRootDSE = GetObject("LDAP://RootDSE")
Set oDomainADSI = GetObject("LDAP://" _
  & oRootDSE.Get ("defaultNamingContext"))
Set oCon = CreateObject("ADODB.Connection") 
Set oCmd = CreateObject("ADODB.Command") 
oCon.Provider = "ADsDSOObject"  
oCon.Open 	
oCmd.ActiveConnection = oCon 	
oCmd.CommandText = "<" & oDomainADSI.ADsPath & ">;" _
	& "(&(objectCategory=person)(objectClass=user)(!cn=*$));distinguishedName,cn,sAMAccountName;" _
	& "subTree"
oCmd.Properties("Page Size") = 100
oCmd.Properties("Timeout") = 30
oCmd.Properties("Cache Results") = False
Set oRst = oCmd.Execute()
On Error Resume Next
oReport.WriteLine("User Name,Never Expires,Expiring Within " _
	& iExpiryNotify & " Days,Expired,Not Expiring")
Do While Not oRst.EOF
	sUser = "LDAP://" & oRst.Fields("distinguishedName")
	Set oUser = GetObject(sUser)
	iCurrentValue = oUser.Get("userAccountControl")
	If iCurrentValue And ADS_UF_DONT_EXPIRE_PASSWD Then
	  iNoExpire = iNoExpire + 1
	  oReport.WriteLine(oRst.Fields("cn") & ",True" & ",,,")
	Else
	  dtmValue = oUser.PasswordLastChanged
	  iTimeInterval = Int(Now - dtmValue)
	  If iTimeInterval >= iMaxPwdAge Then
	      iExpired = iExpired + 1
	      oReport.WriteLine(oRst.Fields("cn") & ",,,True,")
	  Elseif Int((dtmValue + iMaxPwdAge) - Now) <= iExpiryNotify Then
		  iExpiring = iExpiring + 1
		  oReport.WriteLine(oRst.Fields("cn") & ",,True,,")
		  sLog = sLog & oRst.Fields("cn") & " password will " _
		  	  & "expire in " & Int(((dtmValue + iMaxPwdAge) - Now)) & " days" & VbCrLf
	  Else
	  	  iNotExpiring = iNotExpiring + 1
	  	  oReport.WriteLine(oRst.Fields("cn") & ",,,,True")
	  End If
	End If
	iUserNo = iUserNo + 1
	oRst.MoveNext 
Loop
oReport.Close
' sResult = "Total Number of Users: " & iUserNo & VbCrLf & VbCrLf & "Number of accounts with non-expiring passwords: " & vbTab & vbTab & vbTab & iNoExpire & VbCrLf _
' 	& "Number of accounts with passwords expiring within " & iExpiryNotify & " days: " & vbTab & iExpiring & VbCrLf _
' 	& "Number of accounts with expired passwords: " & vbTab & vbTab & vbTab & vbTab & vbTab &  iExpired & VbCrLf _
' 	& "Number of accounts that are not expiring: " & vbTab & vbTab & vbTab & vbTab & vbTab &  iNotExpiring
' WScript.Echo sResult
If sLog = "" Then
	WScript.Quit
Else
	sLog = "Report Generated on " & Now & VbCrLf & VbCrLf & sLog _
		& VbCrLf & VbCrLf & "A report has been generated and saved " _
		& "as " & sReport
	oShell.LogEvent 0, sLog
End If 
