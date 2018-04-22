Option Explicit

Const ForReading = 1
Const ForWriting = 2

Dim sCurrentIPFile, oFSO, url, oHTTP
Dim oFile, sCurrentIP, sPreviousIP
Dim oEmail, sEmailFrom, sEmailTo, sEmailSubject, sSMTPServer

sEmailFrom = "nobodaddy@nowhere.net"
sEmailTo = "stugart@gmail.com"
sEmailSubject = "IP ADDRESS CHANGE"
sSMTPServer = "192.168.0.1"

Set oFSO = CreateObject("Scripting.FileSystemObject")
sCurrentIPFile = "C:\Documents and Settings\Shawn\Desktop\CurrentIP.txt"
url="http://checkip.dyndns.org"

On Error Resume Next

Set oHTTP = CreateObject("MSXML2.XMLHTTP")
Call oHTTP.Open("GET", url, FALSE)
oHTTP.Send
sCurrentIP = (oHTTP.ResponseText)
If Err.Number <> 0 Then
	sCurrentIP = "AN ERROR OCCURRED RETRIEVING YOUR PUBLIC " _
		& "IP ADDRESS FROM " & url & " AT " & Now
End If
sCurrentIP = Replace(sCurrentIP, "<html><head><title>Current IP Check" _
	& "</title></head><body>", "")
sCurrentIP = Replace(sCurrentIP, "</body></html>", "")
sCurrentIP = Left(sCurrentIP, Len(sCurrentIP) - 1) 'for some reason, the url response text 
If Not oFSO.FileExists(sCurrentIPFile) Then		   'has an extra carriage return we'll remove here.	
	Set oFile = oFSO.CreateTextFile(sCurrentIPFile)
	oFile.Write(sCurrentIP)
	oFile.Close
	MailMe sCurrentIP
	WScript.Quit
End If
Set oFile = oFSO.OpenTextFile(sCurrentIPFile, ForReading)
sPreviousIP = oFile.ReadLine

If strComp(sPreviousIP, sCurrentIP, vbTextCompare) = 0 Then
	MailMe sCurrentIP
	WScript.Quit
Else
	Set oFile = oFSO.OpenTextFile(sCurrentIPFile, ForWriting)
	oFile.WriteLine(sCurrentIP)
	MailMe sCurrentIP
End If

Sub MailMe(sNewIP)
	Set oEmail = CreateObject("CDO.Message")
	oEmail.From = sEmailFrom
	oEmail.To = sEmailTo
	oEmail.Subject = sEmailSubject
	oEmail.Textbody = sNewIP
	oEmail.Configuration.Fields.Item _
 	    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	oEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = sSMTPServer
 	oEmail.Configuration.Fields.Item _
 	    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
 	oEmail.Configuration.Fields.Update
	oEmail.Send
End Sub