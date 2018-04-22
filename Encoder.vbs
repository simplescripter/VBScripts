'****************************************************************
'
'For this script to work, you'll need to place the Script Encoder
'file (screnc.exe ) at the root of the C: drive.
'
'
'Last Modified: 12-16-2003
'
'Shawn Stugart
'****************************************************************

Option Explicit
dim oFilesToEncode, oShell, file, strFileOut, i
set oFilesToEncode = WScript.Arguments
set oShell = CreateObject("WScript.Shell")
For i = 0 to oFilesToEncode.Count - 1
    file = oFilesToEncode(i)
    strFileOut = Left(file, Len(file) - 3) & "vbe"
    oShell.Run "%comspec% /c C:\screnc /s " _
        & """" & file & """" & " " & """" & strFileOut & """"
Next