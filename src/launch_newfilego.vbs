set oShell= CreateObject("Wscript.Shell")
set oEnv = oShell.Environment("PROCESS")
oEnv("SEE_MASK_NOZONECHECKS") = 1
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "$$$PATH$$$\launch_newfilego.bat" & Chr(34), 0 'Change $$$PATH$$$ to the path where you are hosting launch_newfilego.bat'
Set WshShell = Nothing
oEnv.Remove("SEE_MASK_NOZONECHECKS")