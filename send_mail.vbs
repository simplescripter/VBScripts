'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: send_mail.vbs
'
' AUTHOR: London , Northwind Traders
' DATE  : 11/25/2003
'
' COMMENT: Example script that sends an email through a remote SMTP server.
'
'==========================================================================

strRemoteSMTP = "Server" ' SMTP server name or IP address
Set objEmail = CreateObject("CDO.Message")
objEmail.From = "yourname@something.com"
objEmail.To = "destination@somewhere.net"
objEmail.Subject = "Your subject here"
objEmail.TextBody = "Enter message string here"
objEmail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
        strRemoteSMTP
objEmail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send