'==============================
' RUN S.A.R.A.H. SERVER
'==============================
Option explicit 

'--------------------------------------
' Main procedure
'--------------------------------------

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess
Dim cJson, objProp
Dim sRunStopProp
Dim sScriptPath, sPluginPath, sRootPath
Dim sNodeJS, sCmdLineNode
Dim bMinServer, bHideServer, bNodeIsRunning
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

If objFSO.FileExists(sRootPath & "WSRNode.cmd") Then
	
	' SARAH Version <= 3 alpha 2 (including 2.8, 2.95 ...)
	sNodeJS           = "WSRNode.cmd"
	sCmdLineNode      = "WSRNode.cmd"
	
ElseIf objFSO.FileExists(sRootPath & "Server_NodeJS.cmd") Then
	
	' SARAH Version >= 3 beta 1 (including 3RC1, 3RC2, 3.0 ...)
	sNodeJS           = "Server_NodeJS.cmd"
	sCmdLineNode      = "Server_NodeJS.cmd"
	
Else
	Wscript.Echo "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If


'-- Read JSon file values

Set cJson    = New VbsJson
sRunStopProp = objFSO.OpenTextFile(sRootPath & "custom.prop").ReadAll
Set objProp  = cJson.Decode(sRunStopProp)
bMinServer   = CBool(objProp("modules")("runstop")("MinimServer"))
bHideServer  = CBool(objProp("modules")("runstop")("HideServer"))

if (bMinServer = "") or (bHideServer = "") then
	Set cJson    = New VbsJson
	sRunStopProp = objFSO.OpenTextFile(sPluginPath & "runstop.prop").ReadAll
	Set objProp  = cJson.Decode(sRunStopProp)
	bMinServer   = CBool(objProp("modules")("runstop")("MinimServer"))
	bHideServer  = CBool(objProp("modules")("runstop")("HideServer"))
end if


'-- Browse processes to know if Server is already running

Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
For Each objProcess in colProcesses
	if not IsNull(objProcess.CommandLine) then		
		' Process for Server NodeJS
		if InStr(1, objProcess.CommandLine, sCmdLineNode) <> 0 then
			bNodeIsRunning = true
		end if
	end if
Next 
Set colProcesses = nothing


'-- Launch server (or not) with its window state = Minimized/Hidden (or not)
	
if bNodeIsRunning = false then
	WshShell.CurrentDirectory = sRootPath
	if bMinServer = true then
		iWindowState = 7
	else
		iWindowState = 1
	end if
	if bHideServer = true then
		iWindowState = 0
	end if
	' Launch Server and wait till it's completely launched
	Return = WshShell.Run(sRootPath & sNodeJS, iWindowState, True)
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