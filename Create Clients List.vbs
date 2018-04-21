'--------------------------------------------------------------------
'*   Creates a Clients.txt file in the root of the C: drive        
'*   containing the names of all systems within the                
'*   subnet that responded to a ping.
'*
'*   Shawn Stugart                              
'*								   
'*   Last Modified 04/14/2004					
'===================================================================

Option Explicit
Dim intIPSubnet, oFSO, oShell, oClientsTxt, i, j, k, oScriptExec1
Dim strPingStdOut1, oScriptExec2, strPingStdOut2, strClient
WScript.StdOut.Write("Enter starting host ID to scan: ")
j = WScript.StdIn.ReadLine
WScript.StdOut.Write("Enter ending host ID to scan: ")
k = WScript.StdIn.ReadLine
WScript.StdOut.Write("Enter network ID (For example, 192.168.1.): ")
intIPSubnet = WScript.StdIn.Readline
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
Set oClientsTxt = oFSO.CreateTextFile("C:\Clients.txt", True)

For i = j To k
    set oScriptExec1 = oShell.Exec("ping -n 1 -w 10 " & intIPSubnet & i)
    WScript.Echo "Pinging " & intIPSubnet & i
    strPingStdOut1 = Lcase(oScriptExec1.StdOut.ReadAll)
    If InStr(strPingStdOut1, "reply from ") Then
        set oScriptExec2 = oShell.Exec("ping -a -n 1 -w 25 " & intIPSubnet & i)
        oScriptExec2.StdOut.SkipLine
        strPingStdOut2 = Lcase(oScriptExec2.StdOut.ReadLine)
        strClient = Split(strPingStdOut2," ",3,1)
        oClientsTxt.WriteLine(strClient(1))        
    End If
Next
    
