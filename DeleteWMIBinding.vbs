sComputer = "192.168.1.100"
sNamespace = "\root\cimv2"
sNameRoot = "NO_Calc"
sInstance = "__FilterToConsumerBinding.Consumer=" _
	& Chr(34) & "\\\\.\\root\\cimv2:ActiveScriptEventConsumer" _
	& ".name=\" & Chr(34) & sNameRoot & "\"& Chr(34) & Chr(34) _
	& ",Filter=" & Chr(34) & "__EventFilter.name=\" & Chr(34) _
	& sNameRoot & "Filter\" & Chr(34) & Chr(34)

'WScript.Echo sInstance
Set oWMI = GetObject("winmgmts:\\" & sComputer & sNamespace)
oWMI.Delete sInstance
