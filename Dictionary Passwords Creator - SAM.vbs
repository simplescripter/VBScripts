'---------------------------------------------------------------------------
'
'   This script loads a text file called c:\dict.txt using the .Dictionary
'   object, uses rnd to pick a random number between 1 and 213558 (the size
'   of this dicitonary), and uses the random number to pick a dictionary
'   password for the 100 users the script generates.
'
'   Shawn Stugart
'===========================================================================


const ForReading = 1
randomize
set oFSO = CreateObject("Scripting.FileSystemObject")
Set oInputFile = oFSO.OpenTextFile("c:\dict.txt", ForReading)
Set oDict = CreateObject("Scripting.Dictionary")

Set oDom = GetObject("WinNT://instructor")
j = 1

Do Until oInputFile.AtEndOfStream
    strText = oInputFile.Readline
    oDict.Add j,strText
    j = j + 1
Loop

For k = 1 to 50
    i = round(rnd * 213558)
    Call CreateUser(k,oDict.Item(i))
Next

WScript.Echo "50 User Created With Passwords"

Sub CreateUser(userNo,password)
    Set oUser = oDom.Create("user","UserNo" & userNo)
    oUser.SetInfo
    oUser.SetPassword password
    oUser.SetInfo
    WScript.Echo oUser.Name & " password set to " & password
End Sub
