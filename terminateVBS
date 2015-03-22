'============ Kill Windows Script Processes ===================
On Error Resume Next 
strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
'======================================================
Set colProcessList = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'Wscript.exe'") 
For Each objProcess in colProcessList 
objProcess.Terminate() 
Next
