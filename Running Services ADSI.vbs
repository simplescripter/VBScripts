'-------------------------------------------------------------
'
'  Example ADSI script that retries a list of running services
'  on the machine specified
'
'  Shawn Stugart
'=============================================================

Option Explicit
dim oComputer, sServices, service
const ADS_SERVICE_RUNNING = &H4
set oComputer = GetObject("WinNT://Santiago")
sServices = "The services running on " & oComputer.ADsPath _
    & " are: " & vbCrLf
For Each Service in oComputer
    If Service.Class = "Service" Then
        If Service.Status = ADS_SERVICE_RUNNING Then
            sServices = sServices & Service.Name & vbCrLf
        End If
    End If
Next
WScript.Echo sServices