Set objShell = WScript.CreateObject("WScript.Shell")
Set objExecObject = objShell.Exec("ipconfig")
Do While Not objExecObject.StdOut.AtEndOfStream
    strText = objExecObject.StdOut.ReadLine()
    If Instr(strText, "IP Address") > 0 Then
        Wscript.Echo strText
    End If
Loop

