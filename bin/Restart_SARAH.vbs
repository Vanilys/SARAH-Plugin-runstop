'==============================
' SCRIPT TO RESTART S.A.R.A.H
'==============================

Dim sScriptPath, sRunSarah, sStopSarah
Dim iReturnValue


iReturnValue = -1

sScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sRunSarah = "Run_SARAH.vbs"
sStopSarah = "Stop_SARAH.vbs"

Set WshShell = WScript.CreateObject("WScript.Shell")

Return = WshShell.Run(sScriptPath & "\" & sStopSarah, 1, true)
Return = WshShell.Run(sScriptPath & "\" & sRunSarah, 1, False)

Set WshShell = nothing


iReturnValue = 0
WScript.Quit(iReturnValue)