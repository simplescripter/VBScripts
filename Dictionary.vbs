'--------------------------------------------------------------------------
'   This script loads a text file called c:\dict.txt using the .Dictionary
'   object, uses rnd to pick a random number between 1 and 213558 (the size
'   of this dicitonary), and displays 100 random dictionary words based 
'   on the number that was chosen
'
'   Shawn Stugart
'==========================================================================
 
const ForReading = 1
randomize
set oFSO = CreateObject("Scripting.FileSystemObject")
Set oInputFile = oFSO.OpenTextFile("c:\dict.txt", ForReading)
Set oDict = CreateObject("Scripting.Dictionary")
j = 1

Do Until oInputFile.AtEndOfStream
    strText = oInputFile.Readline
    oDict.Add j,strText
    j = j + 1
Loop

For k = 1 to 100
    i = round(rnd * 213558)
    WScript.Echo oDict.Item(i)
Next


