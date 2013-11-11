'====================================
' FORCE THE SYSTEM TO RESTART
'====================================

Dim WshShell
Dim sScriptPath

Set WshShell = CreateObject("WScript.Shell")

sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
WshShell.CurrentDirectory = sScriptPath

WshShell.Run "shutdown.exe /r /t 10"

set WshShell = nothing


WScript.Quit(0)