Option Explicit

Dim sComputer, sUser, domainName, dnsName
Dim oFSO, oTS, oShell
'on error resume next 


sComputer= "192.168.1.100"
sUser = "Administrator"
domainName="CLC-D24375B604F" 
dnsName="" 

Set oFSO = CreateObject("Scripting.FileSystemObject") 
Set oTS = oFSO.CreateTextFile("C:\WINDOWS\PCHealth\HelpCtr\Vendors\CN=Microsoft Corporation,L=Redmond,S=Washington,C=US\Remote Assistance\Escalation\Unsolicited\UnSolicitedRCUI.htm") 

oTS.WriteLine "<HTML XMLNS:helpcenter>" 
oTS.WriteLine "<HEAD>" 
oTS.WriteLine "<!--" 
oTS.WriteLine "Copyright (c) 2000 Microsoft Corporation" 
oTS.WriteLine "-->" 
oTS.WriteLine "<helpcenter:context id=idCtx />" 
oTS.WriteLine "<TITLE>Remote Assistance</TITLE>" 
oTS.WriteLine "<!-- The SAF class factory object -->" 
oTS.WriteLine "<OBJECT classid=CLSID:FC7D9E02-3F9E-11d3-93C0-00C04F72DAF7 height=0 id=oSAFClassFactory" 
oTS.WriteLine "width=0></OBJECT>" 
oTS.WriteLine "<script LANGUAGE=" & Chr(34) & "Javascript" & Chr(34) & ">" 
oTS.WriteLine "var g_szRCTicket = null;" 
oTS.WriteLine "var g_szUserName = null;" 
oTS.WriteLine "var g_szDomainName = null;" 
oTS.WriteLine "var g_szSessionId = null;" 
oTS.WriteLine "var g_iExpiry = 5;" 
oTS.WriteLine "var g_oSAFRemoteDesktopConnection = null;" 
oTS.WriteLine "var g_oSAFRemoteConnectionData = null;" 
oTS.WriteLine "var g_oUsersCollection = null;" 
oTS.WriteLine "var g_nUsersLen = null;" 
oTS.WriteLine "var g_oSessionsCollection = null;" 
oTS.WriteLine "var g_nSessionsLen = null;" 
oTS.WriteLine "g_bDebug = false;" 
oTS.WriteLine "function onContinue()" 
oTS.WriteLine "{" 
oTS.WriteLine "var szIncidentFile = null;" 
oTS.WriteLine "var fso = null;" 
oTS.WriteLine "var tempDir = null;" 
oTS.WriteLine "var oInc = null;" 
oTS.WriteLine "g_szDomainName = " & Chr(34) & domainName & Chr(34) & ";" 
oTS.WriteLine "g_szUserName = " & Chr(34) & "" & sUser & "" & Chr(34) & ";" 
oTS.WriteLine "g_szSessionId = -1;" 
oTS.WriteLine "g_oSAFRemoteDesktopConnection = oSAFClassFactory.CreateObject_RemoteDesktopConnection();" 
oTS.WriteLine "g_oSAFRemoteConnectionData = g_oSAFRemoteDesktopConnection.ConnectRemoteDesktop(" & Chr(34) & "" & sComputer & "" & Chr(34) & ");" 
oTS.WriteLine "oInc = oSAFClassFactory.CreateObject_Incident();" 
oTS.WriteLine "oInc.UserName = g_szUserName;" 
oTS.WriteLine "oInc.RCTicketEncrypted = false;" 
oTS.WriteLine "oInc.RcTicket = g_oSAFRemoteConnectionData.ConnectionParms( " & Chr(34) & "" & sComputer & "" & Chr(34) & ", g_szUserName, g_szDomainName, g_szSessionId, " & Chr(34) & "" & Chr(34) & ");" 
oTS.WriteLine "var oDict = oInc.Misc;" 
oTS.WriteLine "var d = new Date();" 
oTS.WriteLine "var iNow = Math.round(Date.parse(d)/1000);" 
oTS.WriteLine "oDict.add(" & Chr(34) & "DtStart" & Chr(34) & ", iNow);" 
oTS.WriteLine "oDict.add(" & Chr(34) & "DtLength" & Chr(34) & ", g_iExpiry);" 
oTS.WriteLine "oDict.add(" & Chr(34) & "IP" & Chr(34) & ", " & Chr(34) & "" & sComputer & "." & dnsName & Chr(34) & ");" 
oTS.WriteLine "oDict.add(" & Chr(34) & "Status" & Chr(34) & ", " & Chr(34) & "Active" & Chr(34) & ");" 
oTS.WriteLine "oDict.add(" & Chr(34) & "URA" & Chr(34) & ", 1);" 
oTS.WriteLine "fso = new ActiveXObject(" & Chr(34) & "Scripting.FileSystemObject" & Chr(34) & ");" 
oTS.WriteLine "tempDir = fso.GetSpecialFolder( 2 );" 
oTS.WriteLine "szIncidentFile = tempDir + " & Chr(34) & "\\UnsolicitedRA" & Chr(34) & " + fso.GetTempName();" 
oTS.WriteLine "oInc.GetXML(szIncidentFile);" 
oTS.WriteLine "var oShell = new ActiveXObject(" & Chr(34) & "WScript.Shell" & Chr(34) & ");" 
oTS.WriteLine "var szRAURL = 'C:\\WINDOWS\\pchealth\\helpctr\\binaries\\helpctr.exe -Mode " & Chr(34) & "hcp://system/Remote Assistance/raura.xml" & Chr(34) & " -url " & Chr(34) & "hcp://system/Remote Assistance/Interaction/Client/RcToolscreen1.htm" & Chr(34) & "' + ' -ExtraArgument " & Chr(34) & "IncidentFile=' + szIncidentFile + '" & Chr(34) & "';" 
oTS.WriteLine "oShell.Run( szRAURL, 1, true );" 
oTS.WriteLine "fso.DeleteFile( szIncidentFile );" 
oTS.WriteLine ("return;") 
oTS.WriteLine ("}") 
oTS.Write "</" 
oTS.Write "SCRIPT" 
oTS.WriteLine ">" 
oTS.WriteLine ("</HEAD>") 
oTS.WriteLine ("<BODY onload=" & Chr(34) & "onContinue();" & Chr(34) & ">") 
oTS.WriteLine ("</BODY>") 
oTS.WriteLine ("</HTML>") 


set oFSO = Nothing     
oTS.Close 
set oTS = Nothing 
Set oShell = CreateObject("WScript.Shell") 
oShell.run "cmd /c start hcp://CN=Microsoft%20Corporation,L=Redmond,S=Washington,C=US/Remote%20Assistance/Escalation/unsolicited/unsolicitedrcui.htm",0,"true" 

