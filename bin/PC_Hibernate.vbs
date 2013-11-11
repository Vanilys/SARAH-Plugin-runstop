'====================================
' HIBERNATE THE SYSTEM
'====================================

Dim WshShell
Dim sScriptPath

Set WshShell = CreateObject("WScript.Shell")

sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
WshShell.CurrentDirectory = sScriptPath

WScript.Sleep(2000)
' Other way :
' oShell.Run "rundll32.exe powrprof.dll,SetSuspendState"
WshShell.Run "shutdown.exe /h"

set WshShell = nothing


WScript.Quit(0)