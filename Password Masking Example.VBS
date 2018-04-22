'---------------------------------------------------
'
' Simple demonstration of the ScriptPW class.  The 
' sole purpose of ScriptPW is to mask data entered
' at the command line.  This script requires that
' the file scriptpw.dll be registered on the system
' (the default for XP and later).
'
'Shawn Stugart
'===================================================

set oUserName = CreateObject("ScriptPW.Password")
WScript.StdOut.Write "Enter Your User Name:"
strUserName = oUserName.GetPassword()
WScript.StdOut.WriteLine

set oPassword = CreateObject("ScriptPW.Password")
WScript.StdOut.Write "Enter Your Password:"
strPassword = oPassword.GetPassword()
WScript.StdOut.WriteLine

WScript.Echo "User is " & strUserName & vbCrLf _
    & "Password is " & strPassword