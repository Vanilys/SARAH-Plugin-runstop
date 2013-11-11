'==============================
' RUN S.A.R.A.H. LOG2CONSOLE
'==============================
Option explicit 

'--------------------------------------
' Main procedure
'--------------------------------------

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess
Dim cJson, objProp
Dim sRunStopProp
Dim sScriptPath, sPluginPath, sRootPath, sConsolePath
Dim sLog2Console, sCmdLineLog2Console
Dim bMinConsole, bConsoleIsRunning
Dim iWindowState
Dim iReturnValue
Dim Return

On Error Resume Next


'-- Initialize parameters
		
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set objFSO = CreateObject( "Scripting.FileSystemObject" )

iReturnValue = -1
sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sPluginPath  = Replace(sScriptPath, "bin\", "")
sRootPath    = Replace(sScriptPath, "plugins\runstop\bin\", "")

includeFile sScriptPath & "Lib_VbsJson.vbs"

sConsolePath = sRootPath & "Log2Console\"
sLog2Console        = "Log2Console.exe"
sCmdLineLog2Console = "Log2Console.exe"


'-- Read JSon file values

Set cJson    = New VbsJson
sRunStopProp = objFSO.OpenTextFile(sRootPath & "custom.prop").ReadAll
Set objProp  = cJson.Decode(sRunStopProp)
bMinConsole  = CBool(objProp("modules")("runstop")("MinimLog2Console"))

if (bMinConsole = "") then
	Set cJson    = New VbsJson
	sRunStopProp = objFSO.OpenTextFile(sPluginPath & "runstop.prop").ReadAll
	Set objProp  = cJson.Decode(sRunStopProp)
	bMinConsole  = CBool(objProp("modules")("runstop")("MinimLog2Console"))
end if


'-- Run the Log2Console application (or not)
	
' Browse processes to know if Log2Console is already running
Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
For Each objProcess in colProcesses
	if not IsNull(objProcess.CommandLine) then		
		' Process for Log2Console
		if InStr(1, objProcess.CommandLine, sCmdLineLog2Console) <> 0 then
			bConsoleIsRunning = true
		end if
	end if
Next 
Set colProcesses = nothing

' If Log2Console isn't already running, then run it !
if bConsoleIsRunning = false then
	WshShell.CurrentDirectory = sConsolePath
	if bMinConsole = true then
		iWindowState = 7
	else
		iWindowState = 1
	end if
	' Launch Client and wait till it's completely launched
	Return = WshShell.Run(sConsolePath & sLog2Console, iWindowState, True)
end if
	
	
'-- Destroy objects

Set objProp = nothing
Set cJson = nothing
Set objWMIService = nothing
Set WshShell = nothing
Set objFSO = nothing


iReturnValue = 0
WScript.Quit(iReturnValue)



'--------------------------------------
' INCLUDE OTHER VBS LIBRARIES
'--------------------------------------

Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub