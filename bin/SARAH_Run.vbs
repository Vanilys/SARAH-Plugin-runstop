'=====================================
' RUN S.A.R.A.H.
'=====================================
Option explicit 

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess
Dim sScriptPath, sRootPath, sPluginPath, sRunServer, sRunClient, sRunLog2Console, sRunActions
Dim cJson, objProp
Dim sRunStopProp
Dim sCmdLineKinect, sCmdLineMicro, sCmdLineClient
Dim bRunServer, bRunClient, bRunLog2Console, bRunActions, bUseKinect, bClientIsRunning
Dim iReturnValue, iTimeBeforeFirstAction, iTimeloop, iTimeOut
Dim TimeBegin
Dim Return

Const DEFAULT_TIMELOOP = 100
Const DEFAULT_TIMEOUT = 30000

On Error Resume Next

'-- Initialize parameters

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set objFSO = CreateObject( "Scripting.FileSystemObject" )

iReturnValue = -1
sScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sPluginPath  = Replace(sScriptPath, "bin\", "")
sRootPath    = Replace(sScriptPath, "plugins\runstop\bin\", "")

sRunServer      = "SARAH_RunServer.vbs"
sRunClient      = "SARAH_RunClient.vbs"
sRunLog2Console = "SARAH_RunLog2Console.vbs"
sRunActions     = "SARAH_RunActions.vbs"

includeFile sScriptPath & "Lib_VbsJson.vbs"

If objFSO.FileExists(sRootPath & "WSRNode.cmd") Then
	
	' SARAH Version <= 3 alpha 2 (including 2.8, 2.95 ...)
	sCmdLineMicro     = "WSRMacro.exe"
	sCmdLineKinect    = "WSRMacro_Kinect.exe"
	
ElseIf objFSO.FileExists(sRootPath & "Server_NodeJS.cmd") Then
	
	' SARAH Version >= 3 beta 1 (including 3RC1, 3RC2, 3.0 ...)
	sCmdLineMicro     = "WSRMacro.exe"
	sCmdLineKinect    = "WSRMacro_Kinect.exe"
	
Else
	Wscript.Echo "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If


'-- Read JSon file values

Set cJson    = New VbsJson
sRunStopProp = objFSO.OpenTextFile(sRootPath & "custom.prop").ReadAll
Set objProp  = cJson.Decode(sRunStopProp)
bRunServer             = CBool(objProp("modules")("runstop")("RunServer"))
bRunClient             = CBool(objProp("modules")("runstop")("RunClient"))
bRunLog2Console        = CBool(objProp("modules")("runstop")("RunLog2Console"))
bRunActions            = CBool(objProp("modules")("runstop")("RunActions"))
bUseKinect             = CBool(objProp("modules")("runstop")("UseKinect"))
iTimeBeforeFirstAction = CInt(objProp("modules")("runstop")("TimeBeforeFirstAction"))

if (bRunServer = "") or (bRunClient = "") or (bRunLog2Console = "")  or (bRunActions = "") or (bUseKinect = "") or (iTimeBeforeFirstAction = "") then
	Set cJson    = New VbsJson
	sRunStopProp = objFSO.OpenTextFile(sPluginPath & "runstop.prop").ReadAll
	Set objProp  = cJson.Decode(sRunStopProp)
	bRunServer             = CBool(objProp("modules")("runstop")("RunServer"))
	bRunClient             = CBool(objProp("modules")("runstop")("RunClient"))
	bRunLog2Console        = CBool(objProp("modules")("runstop")("RunLog2Console"))
	bRunActions            = CBool(objProp("modules")("runstop")("RunActions"))
	bUseKinect             = CBool(objProp("modules")("runstop")("UseKinect"))
	iTimeBeforeFirstAction = CInt(objProp("modules")("runstop")("TimeBeforeFirstAction"))
end if

if bUseKinect  = true then
	sCmdLineClient = sCmdLineKinect
else
	sCmdLineClient = sCmdLineMicro
end if


' -- Launch applications

'-- Run Log2Console
if (bRunLog2Console = true) then 
	Return = WshShell.Run(sScriptPath & sRunLog2Console, 1, false)
end if

'-- Run Server
if (bRunServer = true) then 
	Return = WshShell.Run(sScriptPath & sRunServer, 1, false)
end if 

'-- Run Client
if (bRunClient = true) then 
	Return = WshShell.Run(sScriptPath & sRunClient, 1, false)
end if

'-- Run Actions
if (bRunActions = true) then 

	' Check the client is running
	iTimeloop = DEFAULT_TIMELOOP
	iTimeOut  = DEFAULT_TIMEOUT
	TimeBegin = Time
	Do
		' Wait for 100 ms between each test
		WScript.Sleep(iTimeloop)
		
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

		if DateDiff("s", TimeBegin, Time) > iTimeOut then
			' Exit speaking test if it takes more than 30 sec
			bClientIsRunning = false
			Exit Do
		end if

	Loop Until (bClientIsRunning = true)

	if bClientIsRunning = true then
		' Launch after a few seconds, in order to wait for the client engine
		WScript.Sleep iTimeBeforeFirstAction
		' Run the actions
		Return = WshShell.Run(sScriptPath & sRunActions, 1, true)
	end if
	
end if


'-- Destroy objects

Set objProp = nothing
Set cJson = nothing
Set WshShell = nothing
Set objWMIService = nothing

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