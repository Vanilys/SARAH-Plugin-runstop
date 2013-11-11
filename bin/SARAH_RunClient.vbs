'==============================
' RUN S.A.R.A.H. CLIENT
'==============================
Option explicit 

'--------------------------------------
' Main procedure
'--------------------------------------

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess
Dim cJson, objProp
Dim sRunStopProp
Dim sScriptPath, sPluginPath, sRootPath
Dim sKinect, sMicro, sClient, sCmdLineKinect, sCmdLineMicro, sCmdLineClient
Dim bUseKinect, bUseKinectAudio, bClientIsRunning
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

sAutoItRefresh      = "SystemTray_Refresh.exe"

If objFSO.FileExists(sRootPath & "WSRNode.cmd") Then
	
	' SARAH Version <= 3 alpha 2 (including 2.8, 2.95 ...)
	sMicro            = "WSRMicro.cmd"
	sKinect           = "WSRKinect.cmd"
	sKinectAudio      = "WSRKinect.cmd"
	sCmdLineMicro     = "WSRMacro.exe"
	sCmdLineKinect    = "WSRMacro_Kinect.exe"
	
ElseIf objFSO.FileExists(sRootPath & "Server_NodeJS.cmd") Then
	
	' SARAH Version >= 3 beta 1 (including 3RC1, 3RC2, 3.0 ...)
	sMicro            = "Client_Microphone.cmd"
	sKinect           = "Client_Kinect.cmd"
	sKinectAudio      = "Client_Kinect_Audio.cmd"
	sCmdLineMicro     = "WSRMacro.exe"
	sCmdLineKinect    = "WSRMacro_Kinect.exe"
	
Else
	Wscript.Echo "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If


'-- Read JSon file values

Set cJson       = New VbsJson
sRunStopProp    = objFSO.OpenTextFile(sRootPath & "custom.prop").ReadAll
Set objProp     = cJson.Decode(sRunStopProp)
bUseKinect      = CBool(objProp("modules")("runstop")("UseKinect"))
bUseKinectAudio = CBool(objProp("modules")("runstop")("UseKinectAudio"))

if (bUseKinect = "") or (bUseKinectAudio = "") then
	Set cJson    = New VbsJson
	sRunStopProp = objFSO.OpenTextFile(sPluginPath & "runstop.prop").ReadAll
	Set objProp  = cJson.Decode(sRunStopProp)
	bUseKinect      = CBool(objProp("modules")("runstop")("UseKinect"))
	bUseKinectAudio = CBool(objProp("modules")("runstop")("UseKinectAudio"))
end if


if bUseKinect  = true then
	if bUseKinectAudio = true then
		sClient = sKinectAudio
	else
		sClient = sKinect
	end if
	sCmdLineClient = sCmdLineKinect
else
	sClient = sMicro
	sCmdLineClient = sCmdLineMicro
end if


'-- Browse processes to know if Client is already running

Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
For Each objProcess in colProcesses
	if not IsNull(objProcess.CommandLine) then
		' Process for Client (Kinect or Microphone)
		if InStr(1, objProcess.CommandLine, sCmdLineClient) <> 0 then
			bClientIsRunning = true
		end if
	end if
Next 
Set colProcesses = nothing


'-- Launch client (or not) with its window state = Minimized/Hidden (or not)
	
if bClientIsRunning = false then
	WshShell.CurrentDirectory = sRootPath
	if bMinExe = true then
		iWindowState = 7
	else
		iWindowState = 1
	end if
	if bHideExe = true then
		iWindowState = 0
	end if
	' Refresh the SystemTray icons if the Micro/Kinect has been previously killed, to remove orphaned (ghost) icons
	Return = WshShell.Run(sScriptPath & sAutoItRefresh, 1, true)
	' Launch Client and wait till it's completely launched
	Return = WshShell.Run(sRootPath & sClient, iWindowState, True)
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