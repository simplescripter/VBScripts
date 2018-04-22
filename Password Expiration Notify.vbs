Option Explicit

Const SEC_IN_DAY = 86400
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000

Dim iExpiryNotify, sDomain, oDomain, iMaxPwdAge
Dim oUser, iCurrentValue, iNoExpire, iExpiring
Dim sResult, dtmValue, iTimeInterval, iExpired, sLog
Dim sExpires, oRootDSE, oDomainADSI, oCon, oCmd, oShell, sReport
Dim oRst, oFSO, oReport, iDays, sUser, iUserNo, iNotExpiring
Dim iSendMail, iCreateExcelReport, sEmailFrom, oEmail, sEmailSubject
Dim sSMTPServer


'******************************************************************************
'	Set variables here

iExpiryNotify = int(410) 'number of days before password expiration to notify
sDomain = "fourthcoffee" 'set domain name here
iSendMail = True 'set to True to send email to users with expiring passwords 
sSMTPServer = "LON-DC1.Fourthcoffee.com"
sEmailFrom = "test@fourthcoffee.com"
sEmailSubject = "Password change requested"
iCreateExcelReport = False

'******************************************************************************
iNoExpire = 0
iExpiring = 0
iExpired = 0
iNotExpiring = 0

Set oShell = CreateObject("WScript.Shell")
Set oDomain = GetObject("WinNT://" & sDomain)
Set oEmail = CreateObject("CDO.Message")

iMaxPwdAge = oDomain.Get("MaxPasswordAge")
iMaxPwdAge = (iMaxPwdAge/SEC_IN_DAY)

If iCreateExcelReport = True Then
	sReport = "C:\PasswordReport-" & FormatDateTime(Date,vbLongDate) & ".csv"
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oReport = oFSO.CreateTextFile(sReport,True)
	oReport.WriteLine("User Name,Never Expires,Expiring Within " _
	& iExpiryNotify & " Days,Expired,Not Expiring")
End If

Set oRootDSE = GetObject("LDAP://RootDSE")
Set oDomainADSI = GetObject("LDAP://" _
  & oRootDSE.Get ("defaultNamingContext"))
Set oCon = CreateObject("ADODB.Connection") 
Set oCmd = CreateObject("ADODB.Command") 
oCon.Provider = "ADsDSOObject"  
oCon.Open 	
oCmd.ActiveConnection = oCon 	
oCmd.CommandText = "<" & oDomainADSI.ADsPath & ">;" _
	& "(&(objectCategory=person)(objectClass=user)(!cn=*$));distinguishedName,cn,sAMAccountName,mail;" _
	& "subTree"
oCmd.Properties("Page Size") = 100
oCmd.Properties("Timeout") = 30
oCmd.Properties("Cache Results") = False
Set oRst = oCmd.Execute()
On Error Resume Next

Do While Not oRst.EOF
	sUser = "LDAP://" & oRst.Fields("distinguishedName")
	Set oUser = GetObject(sUser)
	iCurrentValue = oUser.Get("userAccountControl")
	If iCurrentValue And ADS_UF_DONT_EXPIRE_PASSWD Then
	  iNoExpire = iNoExpire + 1
	  If iCreateExcelReport = True Then
	  	oReport.WriteLine(oRst.Fields("cn") & ",True" & ",,,")
	  End If
	Else
	  dtmValue = oUser.PasswordLastChanged
	  iTimeInterval = Int(Now - dtmValue)
	  If iTimeInterval >= iMaxPwdAge Then
	      iExpired = iExpired + 1
	      If iCreateExcelReport = True Then
	      	oReport.WriteLine(oRst.Fields("cn") & ",,,True,")
	      End If
	  Elseif Int((dtmValue + iMaxPwdAge) - Now) <= iExpiryNotify Then
		  iExpiring = iExpiring + 1
		  If iCreateExcelReport = True Then
		  	oReport.WriteLine(oRst.Fields("cn") & ",,True,,")
		  End If
		  sLog = sLog & oRst.Fields("cn") & " password will " _
		  	  & "expire in " & Int(((dtmValue + iMaxPwdAge) - Now)) & " days" & VbCrLf
		  If iSendMail = True Then
		      Call Mailer(oRst.Fields("mail"), oRst.Fields("cn"), Int(((dtmValue + iMaxPwdAge) - Now)))
		  End If
	  Else
	  	  iNotExpiring = iNotExpiring + 1
	  	  If iCreateExcelReport = True Then
	  	  	oReport.WriteLine(oRst.Fields("cn") & ",,,,True")
	  	  End If
	  End If
	End If
	iUserNo = iUserNo + 1
	oRst.MoveNext 
Loop
oReport.Close
If sLog = "" Then
	WScript.Quit
Else
	sLog = "Report Generated on " & Now & VbCrLf & VbCrLf & sLog _
		& VbCrLf & VbCrLf & "A report has been generated and saved " _
		& "as " & sReport
	oShell.LogEvent 0, sLog
End If 

Sub Mailer(sEmail, sUser, iExpiringIn)
	oEmail.From = sEmailFrom
	oEmail.To = sEmail
	oEmail.Subject = sEmailSubject
	oEmail.Textbody = sUser & ", your password will be expiring in " & iExpiringIn _
		& " days.  Please log on and reset it now."
	oEmail.Configuration.Fields.Item _
    	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	oEmail.Configuration.Fields.Item _
    	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = sSMTPServer
	oEmail.Configuration.Fields.Item _
    	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	oEmail.Configuration.Fields.Update
	oEmail.Send
End Sub