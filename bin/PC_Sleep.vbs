'====================================
' SLEEP THE SYSTEM
'====================================

' WIZMO is a powerful tool to sleep/hibernate and so on
' Made by Steve Gibson
' SITE : https://www.grc.com/wizmo/wizmo.htm

Dim WshShell
Dim sScriptPath

Set WshShell = CreateObject("WScript.Shell")

sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
WshShell.CurrentDirectory = sScriptPath

WshShell.Run sScriptPath & "wizmo quiet standby"

set WshShell = nothing


WScript.Quit(0)