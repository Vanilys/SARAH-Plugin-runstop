'==============================
' STOP S.A.R.A.H. LOG2CONSOLE
'==============================
Option explicit 

'--------------------------------------
' Main procedure
'--------------------------------------

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess
Dim cJson, objProp
Dim sRunStopProp
Dim sScriptPath, sPluginPath, sRootPath, sCmdLineLog2Console, sAutoItExit, sAutoItRefresh, sExitTrayIcon_Console
Dim bStopGracefully
Dim iReturnValue, iTimeBeforeKill, iNbProcessRunning
Dim TimeBegin
Dim Return


'-- Initialize parameters

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set objFSO = CreateObject( "Scripting.FileSystemObject" )

iReturnValue = -1
sScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sPluginPath = Replace(sScriptPath, "bin\", "")
sRootPath   = Replace(sScriptPath, "plugins\runstop\bin\", "")

includeFile sScriptPath & "Lib_VbsJson.vbs"

sAutoItExit           = "SystemTray_Exit.exe"
sAutoItRefresh        = "SystemTray_Refresh.exe"
sExitTrayIcon_Console = sAutoItExit & " /Console"
sCmdLineLog2Console   = "Log2Console.exe"


'-- Read JSon file values

Set cJson = New VbsJson
sRunStopProp    = objFSO.OpenTextFile(sRootPath & "custom.prop").ReadAll
Set objProp = cJson.Decode(sRunStopProp)
bStopGracefully     = CBool(objProp("modules")("runstop")("StopGracefully"))
iTimeBeforeKill     = CInt(objProp("modules")("runstop")("TimeBeforeKill"))

if (bStopGracefully = "") or (iTimeBeforeKill = "") then
	Set cJson    = New VbsJson
	sRunStopProp = objFSO.OpenTextFile(sPluginPath & "runstop.prop").ReadAll
	Set objProp  = cJson.Decode(sRunStopProp)
	bStopGracefully     = CBool(objProp("modules")("runstop")("StopGracefully"))
	iTimeBeforeKill     = CInt(objProp("modules")("runstop")("TimeBeforeKill"))
end if

iTimeBeforeKill = Round(iTimeBeforeKill / 1000,0)
if iTimeBeforeKill > 1000 then iTimeBeforeKill = 60


'-- Browse all of the process, and check their command line

WshShell.CurrentDirectory = sScriptPath
Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
For Each objProcess in colProcesses
	if not IsNull(objProcess.CommandLine) then
		
		' Process for the console
		if InStr(1, objProcess.CommandLine, sCmdLineLog2Console) <> 0 then
			'msgbox objProcess.ProcessId & " | " & objProcess.ParentProcessId & " | " & objProcess.CommandLine	
			' WARNING
			' FINALLY I PREFER TO KILL THE PROCESS INSTEAD OF TRYING TO CLOSE IT GRACEFULLY
			'if bStopGracefully = false then 
				objProcess.Terminate()
			'else
			'	Return = WshShell.Run(sScriptPath & sExitTrayIcon_Console, 1, true)
			'end if
		end if
		
	end if
Next


'-- Check if application are still running
'   If stopping gracefully, it can take a few seconds before all application are really closed
'   => If the time is too long (> TimeOut), kill process

if bStopGracefully = false then 
	iReturnValue = 0
else
	TimeBegin = Time
	Do  
		iNbProcessRunning = 0
		Set colProcesses = nothing
		Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
		
		For Each objProcess in colProcesses
			if not IsNull(objProcess.CommandLine) then
								
				' Console
				if InStr(1, objProcess.CommandLine, sCmdLineLog2Console) <> 0 then
					iNbProcessRunning = iNbProcessRunning + 1
					if DateDiff("s", TimeBegin, Time) >= iTimeBeforeKill then
						objProcess.Terminate()
					end if
				end if
				
			end if
		Next
		WScript.Sleep(50)
	Loop Until iNbProcessRunning=0
	iReturnValue = 0
end if


'-- Refresh the SystemTray icons if the Log2Console has been killed, to remove orphaned (ghost) icons

Return = WshShell.Run(sScriptPath & sAutoItRefresh, 1, true)

'--  Destroy objects

Set objProp = nothing
Set cJson = nothing
Set colProcesses = nothing
Set objWMIService = nothing
Set WshShell = nothing
Set objFSO = nothing


WScript.Quit(iReturnValue)



'--------------------------------------
' INCLUDE OTHER VBS LIBRARIES
'--------------------------------------

Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub