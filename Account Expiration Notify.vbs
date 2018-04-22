Option Explicit

Dim iExpiryNotify, sDomain, oDomain, dtmAccountExpiration
Dim oUser, iCurrentValue, sExpirationResult
Dim sResult, dtmValue, iTimeInterval, sLog
Dim sExpires, oRootDSE, oDomainADSI, oCon, oCmd, oShell, sReport
Dim oRst, oFSO, oReport, iDays, sUser, iUserNo, iNotExpiring
Dim iSendMail, iCreateExcelReport, sEmailFrom, oEmail, sEmailSubject
Dim sSMTPServer, iDaysRemaining


'******************************************************************************
'	Set variables here

iExpiryNotify = 10 'number of days before account expiration to notify
sDomain = "fourthcoffee" 'set domain name here
iSendMail = False 'set to True to send email to users with expiring accounts 
sSMTPServer = "LON-DC1.Fourthcoffee.com"
sEmailFrom = "test@fourthcoffee.com"
sEmailSubject = "Your user account is expiring"

'******************************************************************************
On Error Resume Next

Set oEmail = CreateObject("CDO.Message")
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

Do While Not oRst.EOF
	sUser = "LDAP://" & oRst.Fields("distinguishedName")
	Set oUser = GetObject(sUser)
	dtmAccountExpiration = oUser.AccountExpirationDate
	If Err.Number = -2147467259 Or dtmAccountExpiration = "1/1/1970" Then
		Err.Clear
	Else
		iDaysRemaining = DateDiff("d",Date,dtmAccountExpiration)
		If int(iDaysRemaining) <= int(iExpiryNotify) Then
			If iSendMail = True Then
				Mailer oRst.Fields("mail"), oRst.Fields("cn"), iDaysRemaining
			End If
			sLog = sLog & "Account " & oRst.Fields("cn") & " will expire on " & dtmAccountExpiration _
				 & ": " & iDaysRemaining & " days remaining." & vbCrLf
		End If
	End If
	oRst.MoveNext
Loop	

If sLog = "" Then
	WScript.Echo "Complete.  No accounts are expiring in the next " & iExpiryNotify & " days."
Else	
	WScript.Echo sLog
End If

Sub Mailer(sEmail, sUser, iExpiringIn)
	oEmail.From = sEmailFrom
	oEmail.To = sEmail
	oEmail.Subject = sEmailSubject
	oEmail.Textbody = sUser & ", your account will be expiring in " & iExpiringIn _
		& " days.  Your access will be revoked after that time." & vbCrLf & vbCrLf _
		& "Thank you."
	oEmail.Configuration.Fields.Item _
    	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	oEmail.Configuration.Fields.Item _
    	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = sSMTPServer
	oEmail.Configuration.Fields.Item _
    	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	oEmail.Configuration.Fields.Update
	oEmail.Send
End Sub
	