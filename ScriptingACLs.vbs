' Code adapted from the Scripting Guys at Microsoft, 
' http://www.microsoft.com/technet/technetmag/issues/2006/05/ScriptingGuy/default.aspx

sComputer = "."
sFolderOrFile = "c:\"
Set oWMI = GetObject("winmgmts:\\" & sComputer & "\root\cimv2")
Set oFile = oWMI.Get("Win32_LogicalFileSecuritySetting='" & sFolderOrFile & "'")

If oFile.GetSecurityDescriptor(oSD) = 0 Then

Wscript.Echo "Owner: " & oSD.Owner.Name
Wscript.Echo

For Each oACE in oSD.DACL
    Wscript.Echo "Trustee: " & oACE.Trustee.Domain & "\" & oACE.Trustee.Name

    If oACE.AceType = 0 Then
        sACEType = "Allowed"
    Else
        sACEType = "Denied"
    End If
    Wscript.Echo "Ace Type: " & sACEType

    Wscript.Echo "Ace Flags:"

    If oACE.AceFlags AND 1 Then
        Wscript.Echo vbTab & "Child objects that are not containers inherit permissions."
    End If

    If oACE.AceFlags AND 2 Then
        Wscript.Echo vbTab & "Child objects inherit and pass on permissions."
    End If

    If oACE.AceFlags AND 4 Then
        Wscript.Echo vbTab & "Child objects inherit but do not pass on permissions."
    End If

    If oACE.AceFlags AND 8 Then
        Wscript.Echo vbTab & "Object is not affected by but passes on permissions."
    End If

    If oACE.AceFlags AND 16 Then
        Wscript.Echo vbTab & "Permissions have been inherited."
    End If

    Wscript.Echo "Access Masks:"
    If oACE.AccessMask AND 1048576 Then
        Wscript.Echo vbtab & "Synchronize"
    End If

    If oACE.AccessMask AND 524288 Then
        Wscript.Echo vbtab & "Write owner"
    End If
    If oACE.AccessMask AND 262144 Then
        Wscript.Echo vbtab & "Write ACL"
    End If
    If oACE.AccessMask AND 131072 Then
        Wscript.Echo vbtab & "Read security"
    End If
    If oACE.AccessMask AND 65536 Then
        Wscript.Echo vbtab & "Delete"
    End If
    If oACE.AccessMask AND 256 Then
        Wscript.Echo vbtab & "Write attributes"
    End If
    If oACE.AccessMask AND 128 Then
        Wscript.Echo vbtab & "Read attributes"
    End If
    If oACE.AccessMask AND 64 Then
        Wscript.Echo vbtab & "Delete dir"
    End If
    If oACE.AccessMask AND 32 Then
        Wscript.Echo vbtab & "Execute"
    End If
    If oACE.AccessMask AND 16 Then
        Wscript.Echo vbtab & "Write extended attributes"
    End If
    If oACE.AccessMask AND 8 Then
        Wscript.Echo vbtab & "Read extended attributes"
    End If
    If oACE.AccessMask AND 4 Then
        Wscript.Echo vbtab & "Append"
    End If

    If oACE.AccessMask AND 2 Then
        Wscript.Echo vbtab & "Write"
    End If

    If oACE.AccessMask AND 1 Then
        Wscript.Echo vbtab & "Read"
    End If

    Wscript.Echo
    Wscript.Echo
Next

End If
