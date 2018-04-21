'------------------------------------------------------
'
'  Example Code to verify whether or not the script is
'  executing under cscript.exe
'
'  Shawn Stugart
'======================================================

hostname = lcase(right(WSCript.Fullname, 11))
If hostname = "wscript.exe" Then
    WScript.Echo "This script requires cscript.exe"
    WScript.Quit
End If

WScript.Echo "You're running " & WScript.Fullname
WScript.Echo "Translated, you're running " & hostname