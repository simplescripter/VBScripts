'USB Storage Monitor
'
'Vinicius Canto
'MVP Visual Developer - Scripting
'
'Grupo PET Computação - Universidade de São Paulo - Brasil
'http://viniciuscanto.blogspot.com

'Disabling error messages...
On Error Resume Next


'Main routine
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery("SELECT * FROM __InstanceCreationEvent WITHIN 10 WHERE Targetinstance ISA 'Win32_PNPEntity' and TargetInstance.DeviceId like '%USBStor%'")
Do
Set objLatestEvent = colMonitoredEvents.NextEvent
Notifier(objLatestEvent.TargetInstance)
Loop

Sub Notifier(object)
Set objNet = CreateObject("Wscript.Network")

'You can change the function below to perform other actions
SendMailWithoutSSL _
"admin@network.com", _
"USB storage detected on " & objNet.Computername, _
"machine@network.com", _
"The user " & objNet.Username & " connected an USB Storage device on computer.", _
"smtp.network.com", _
25, _
"user", _
"pass"
End Sub




' CDOSYS official documentation:
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wss/wss/_cdo_queue_top.asp
'
' by Vinicius Canto
Sub SendMailWithoutSSL(strDestination, strTitle, strFrom, strMessage, strSMTP, intPort, strUsername, strPassword)
set oMessage = CreateObject("CDO.Message")
set oConf = CreateObject("CDO.Configuration")
Set oFields = oConf.Fields



oFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTP
oFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = intPort
oFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic: Auth with user and password sent with plain text
oFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strUsername
oFields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strPassword
oFields.Item("http://schemas.microsoft.com/cdo/configuration/Smtpusessl") = false
oFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1: Using local SMTP; 2: Using port; 3: Using Exchange
oFields.Update

oMessage.Fields.Item("urn:schemas:mailheader:to") = strDestination
oMessage.Fields.Item("urn:schemas:mailheader:from") = strFrom
oMessage.Fields.Item("urn:schemas:mailheader:sender") = strFrom 'reply-to
oMessage.Fields.Item("urn:schemas:mailheader:subject")= strTitle
oMessage.Fields.Item("urn:schemas:mailheader:x-mailer") = "Vinicius Small Mail System -- by Vinicius Canto "
oMessage.Fields.Update

oMessage.Configuration = oConf

oMessage.TextBody = strMessage
oMessage.Send
End Sub

