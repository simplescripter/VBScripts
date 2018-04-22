Option Explicit

dim oFSO, strText, regEx, Match, Matches, j, dtmNow, oShell, strLogPath
dim oInputFile, RetStr, strLogFile, sPattern, strDB, oScriptExec, strScriptName

const ForReading = 1
Const EVENT_INFORMATION = 4
Const EVENT_ERROR = 1
dtmNow = FormatDateTime(Date,vbLongDate)
strScriptName = WScript.ScriptName
set oFSO = CreateObject("Scripting.FileSystemObject")
set oShell = CreateObject("WScript.Shell")
strLogPath = "C:\" & dtmNow & "-" & "sec-template-baseline.log"
set strLogFile = oFSO.CreateTextFile(strLogPath, True)
strDB = "C:\sec-template-baseline.sdb"
Set oScriptExec = oShell.Exec("secedit /analyze /db " _
    & strDB & " /log " & """" & strLogPath & """" & " /verbose")
strLogFile.Close
WScript.Sleep 3000

Set oInputFile = oFSO.OpenTextFile(strLogPath, ForReading,, -1)

sPattern = "\Mismatch"

If Err.Number <> 0 Then
    WScript.Echo "Could not open " & strLogFile
    WScript.Quit
End If

Set regEx = New RegExp   ' Create a regular expression.
regEx.Pattern = sPattern   ' Set pattern.
regEx.IgnoreCase = True   ' Set case insensitivity.
regEx.Global = True   ' Set global applicability.

Do Until oInputFile.AtEndOfStream
    strText = oInputFile.Readline
    RegExpTest sPattern, strText
Loop

Function RegExpTest(patrn, strng)
   Set Matches = regEx.Execute(strText)   ' Execute search.
   For Each Match in Matches   ' Iterate Matches collection.
      j = j + 1
      RetStr = RetStr & strng & "'." & vbCRLF
   Next
   RegExpTest = RetStr & vbCrLf & vbCrlF & j & " Mis-matches found."
End Function

If j > 0 Then
   oShell.LogEvent EVENT_ERROR, strScriptName & vbCrLf _
      & vbCrLf & RegExpTest("is.", "IS1")
   MsgBox(RegExpTest("is.", "IS1"))
Else 
   oShell.LogEvent EVENT_INFORMATION, strScriptName & vbCrLf _
      & vbCrLf & "No mismatches were found."
   MsgBox "No mismatches were found when comparing your local security to " _
      & "the " & strDB & " security database."
End If


MsgBox "This information has been written to the Application Log."


