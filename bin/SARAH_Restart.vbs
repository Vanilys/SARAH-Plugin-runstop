'=====================================
' RESTART (STOP THEN START) S.A.R.A.H.
'=====================================
Option explicit 

Dim WshShell
Dim sScriptPath, sRunSarah, sStopSarah
Dim iReturnValue
Dim Return


iReturnValue = -1

sScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sRunSarah = "SARAH_Run.vbs"
sStopSarah = "SARAH_Stop.vbs"

Set WshShell = WScript.CreateObject("WScript.Shell")

' Stop SARAH, and wait until the processes are finished
Return = WshShell.Run(sScriptPath & sStopSarah, 1, true)

' Run SARAH
Return = WshShell.Run(sScriptPath & sRunSarah, 1, False)

Set WshShell = nothing


iReturnValue = 0
WScript.Quit(iReturnValue)