'--------------------------------------------------------------------------
'
'   Example code to send an email through a defined SMTP server.  This 
'   extra information is necessary if a local SMTP service is not running
'   on the script client or if the client is not MAPI-configured already.
'
'   Shawn Stugart
'==========================================================================

Set objEmail = CreateObject("CDO.Message")
objEmail.From = "someone@somewhere.net"
objEmail.To = "nobody@nowhere.com"
objEmail.Subject = "Server down"
objEmail.Textbody = "Server1 is no longer accessible over the network."
objEmail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
        "192.168.0.1"
objEmail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send
