'SendMail.vbs
'Published 2/27/01 - Windows Script Community, MSN.com

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set appOutl = Wscript.CreateObject("Outlook.Application")

'Set a reference to the MailItem object.
Set maiMail = appOutl.CreateItem(0)
'Get an address from the user.
maiMail.Recipients.Add(InputBox("Enter name of message recipient"))
'Add subject and body text.
maiMail.Subject = "Testing mail by Automation"
maiMail.Body = "Message body text"
        
'Send the mail.
maiMail.Send
    
'Close object references.
Set appOutl = Nothing
Set maiMail = Nothing
Set recMessage = Nothing
