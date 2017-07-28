Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c intercept.exe /apply"
oShell.Run strArgs, 0, false