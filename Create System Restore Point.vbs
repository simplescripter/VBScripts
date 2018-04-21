strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\default") 
Set objItem = objWMIService.Get("SystemRestore") 
errResults = objItem.CreateRestorePoint _ 
    ("Scripted restore point on " & Now, 0, 100)
