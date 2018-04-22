Set objArgs = WScript.Arguments.Named
strUser = LCase(objArgs.Item("u"))
strComputer = LCase(objArgs.Item("c"))
strDomain = LCase(objArgs.Item("d"))

CheckForCScript

If WScript.Arguments.Named.Exists("u") Then
    WScript.StdOut.WriteLine "User Account is " & strUser
Else
    WScript.StdOut.WriteLine "You must provide the /u: parameter."
    WScript.Quit
End If
If WScript.Arguments.Named.Exists("c") Then
    compArray = Split(strComputer,",")
    For i = 0 to UBound(compArray)
        WScript.StdOut.WriteLine "Computer Name is " & compArray(i)
    Next
End If
If WScript.Arguments.Named.Exists("d") Then
    WScript.StdOut.WriteLine "Domain is " & strDomain
End If


Sub CheckForCScript()
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    WScript.Quit
	End If
End Sub