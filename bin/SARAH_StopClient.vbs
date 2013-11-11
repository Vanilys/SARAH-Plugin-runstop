'==============================
' STOP S.A.R.A.H. CLIENT
'==============================
Option explicit 

'--------------------------------------
' Main procedure
'--------------------------------------

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess
Dim cJson, objProp
Dim sRunStopProp
Dim sScriptPath, sPluginPath, sRootPath, sCmdLineMicro, sCmdLineKinect, sAutoItExit, sAutoItRefresh, sExitTrayIcon_Speech
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
sExitTrayIcon_Speech  = sAutoItExit & " /Speech"

If objFSO.FileExists(sRootPath & "WSRNode.cmd") Then
	
	' SARAH Version <= 3 alpha 2 (including 2.8, 2.95 ...)
	sCmdLineMicro       = "WSRMacro.exe"
	sCmdLineKinect      = "WSRMacro_Kinect.exe"
	
ElseIf objFSO.FileExists(sRootPath & "Server_NodeJS.cmd") Then
	
	' SARAH Version >= 3 beta 1 (including 3RC1, 3RC2, 3.0 ...)
	sCmdLineMicro       = "WSRMacro.exe"
	sCmdLineKinect      = "WSRMacro_Kinect.exe"
	
Else
	Wscript.Echo "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If


'-- Read JSon file values

Set cJson = New VbsJson
sRunStopProp    = objFSO.OpenTextFile(sRootPath & "custom.prop").ReadAll
Set objProp = cJson.Decode(sRunStopProp)
bStopGracefully = CBool(objProp("modules")("runstop")("StopGracefully"))
iTimeBeforeKill = CInt(objProp("modules")("runstop")("TimeBeforeKill"))

if (bStopGracefully = "") or (iTimeBeforeKill = "") then
	Set cJson    = New VbsJson
	sRunStopProp = objFSO.OpenTextFile(sPluginPath & "runstop.prop").ReadAll
	Set objProp  = cJson.Decode(sRunStopProp)
	bStopGracefully = CBool(objProp("modules")("runstop")("StopGracefully"))
	iTimeBeforeKill = CInt(objProp("modules")("runstop")("TimeBeforeKill"))
end if

iTimeBeforeKill = Round(iTimeBeforeKill / 1000, 0)
if iTimeBeforeKill > 1000 then iTimeBeforeKill = 60


'-- Browse all of the process, and check their command line

WshShell.CurrentDirectory = sScriptPath
Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
For Each objProcess in colProcesses
	if not IsNull(objProcess.CommandLine) then
			
		' Process for standard Microphone
		if InStr(1, objProcess.CommandLine, sCmdLineMicro) <> 0 then
			'msgbox objProcess.ProcessId & " | " & objProcess.ParentProcessId & " | " & objProcess.CommandLine			
			if bStopGracefully = false then 				
				objProcess.Terminate()
			else
				Return = WshShell.Run(sScriptPath & sExitTrayIcon_Speech, 1, true)
			end if
		end if
		
		' Process for Kinect
		if InStr(1, objProcess.CommandLine, sCmdLineKinect) <> 0 then
			'msgbox objProcess.ProcessId & " | " & objProcess.ParentProcessId & " | " & objProcess.CommandLine				
			if bStopGracefully = false then 
				objProcess.Terminate()
			else
				Return = WshShell.Run(sScriptPath & sExitTrayIcon_Speech, 1, true)
			end if
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
				
				' Standard microphone
				if InStr(1, objProcess.CommandLine, sCmdLineMicro) <> 0 then
					iNbProcessRunning = iNbProcessRunning + 1
					if DateDiff("s", TimeBegin, Time) >= iTimeBeforeKill then
						objProcess.Terminate()
					end if
				end if
				
				' Kinect
				if InStr(1, objProcess.CommandLine, sCmdLineKinect) <> 0 then
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


'-- Refresh the SystemTray icons if the Micro/Kinect has been killed, to remove orphaned (ghost) icons

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