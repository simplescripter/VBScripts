'--------------------------------------------------------------------
'*   Creates a Hosts file containing the names and IPs of all       
'*     systems within the subnet that responded to a ping.          
'*
'*   Shawn Stugart
'*						    		    
'*   Last Modified 04/14/2004					    
'====================================================================

Option Explicit
Dim intIPSubnet, oFSO, oShell, oHostsTxt, i, oScriptExec1
Dim strPingStdOut1, oScriptExec2, strPingStdOut2, strClient
intIPSubnet = "192.168.25."  ' come back and replace with a cscript prompt
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
Set oHostsTxt = oFSO.CreateTextFile("c:\windows\system32\drivers\etc\hosts", True)

For i = 100 To 200
    set oScriptExec1 = oShell.Exec("ping -n 1 -w 10 " & intIPSubnet & i)
    WScript.Echo "Pinging " & intIPSubnet & i
    strPingStdOut1 = Lcase(oScriptExec1.StdOut.ReadAll)
    If InStr(strPingStdOut1, "reply from ") Then
        set oScriptExec2 = oShell.Exec("ping -a -n 1 -w 25 " & intIPSubnet & i)
        oScriptExec2.StdOut.SkipLine
        strPingStdOut2 = Lcase(oScriptExec2.StdOut.ReadLine)
        strClient = Split(strPingStdOut2," ",3,1)
        oHostsTxt.WriteLine(intIPSubNet & i & vbTab & strClient(1))        
    End If
Next
    
