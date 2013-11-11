'==============================
' SCRIPT TO RUN S.A.R.A.H
'==============================

'--------------------------------------
' Main procedure
'--------------------------------------

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess
Dim cJson, objProp
Dim sRunStopProp
Dim sScriptPath, sPluginPath, sRootPath, sConsolePath, sIsSpeakingLoop
Dim sNodeJS, sKinect, sMicro, sLog2Console, sClient
Dim sCmdLineNode, sCmdLineKinect, sCmdLineMicro, sCmdLineConhost, sCmdLineLog2Console, sCmdLineClient
Dim iWindowState, iTimeBeforeFirstAction, iTimeBetweenEachAction
Dim bUseKinect, bUseKinectAudio, bMinExe, bHideExe, bRunConsole, bMinConsole, bConsoleIsRunning, bNodeIsRunning, bClientIsRunning
Dim sURL, sToSend
Dim iReturnValue, IsSpeaking, iAction
Dim sActions, sCurrAction, sActionNum, sActionURL
Dim aActions
Dim bIsPlugin, bIsSentence, bIsExe, bIsEmulation
Dim TimeBegin

Const DEF_TIMELOOP = 200
Const DEF_TIMEOUT = 30000

'-- Initialize parameters
		
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set objFSO = CreateObject( "Scripting.FileSystemObject" )

iReturnValue = -1
sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sPluginPath  = Replace(sScriptPath, "bin\", "")
sRootPath    = Replace(sScriptPath, "plugins\runstop\bin\", "")

includeFile sScriptPath & "Lib_VbsReadIni.vbs"
includeFile sScriptPath & "Lib_VbsJson.vbs"
includeFile sScriptPath & "Lib_VbsSendHTTP.vbs"

sConsolePath = sRootPath & "Log2Console\"
sIsSpeakingLoop     = "IsSpeaking_LoopTest.vbs"
sAutoItRefresh      = "SystemTray_Refresh.exe"
sLog2Console        = "Log2Console.exe"
sCmdLineLog2Console = "Log2Console.exe"

If objFSO.FileExists(sRootPath & "WSRNode.cmd") Then
	
	' SARAH Version <= 3 alpha 2 (including 2.8, 2.95 ...)
	sNodeJS           = "WSRNode.cmd"
	sMicro            = "WSRMicro.cmd"
	sKinect           = "WSRKinect.cmd"
	sKinectAudio      = "WSRKinect.cmd"
	sCmdLineNode      = "WSRNode.cmd"
	sCmdLineMicro     = "WSRMacro.exe"
	sCmdLineKinect    = "WSRMacro_Kinect.exe"
	
ElseIf objFSO.FileExists(sRootPath & "Server_NodeJS.cmd") Then
	
	' SARAH Version >= 3 beta 1 (including 3RC1, 3RC2, 3.0 ...)
	sNodeJS           = "Server_NodeJS.cmd"
	sMicro            = "Client_Microphone.cmd"
	sKinect           = "Client_Kinect.cmd"
	sKinectAudio      = "Client_Kinect_Audio.cmd"
	sCmdLineNode      = "Server_NodeJS.cmd"
	sCmdLineMicro     = "WSRMacro.exe"
	sCmdLineKinect    = "WSRMacro_Kinect.exe"
	
Else
	Wscript.Echo "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If


'-- Read ini file values

'bUseKinect             = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "UseKinect"))
'bUseKinectAudio        = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "UseKinectAudio"))
'bMinExe                = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "MinimExe"))
'bHideExe               = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "HideExe"))
'bRunConsole            = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "RunLog2Console"))
'bMinConsole            = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "MinimLog2Console"))
'iTimeBeforeFirstAction = CInt(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN_ACTIONS", "TimeBeforeFirstAction"))
'iTimeBetweenEachAction = CInt(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN_ACTIONS", "TimeBetweenEachAction"))

Set cJson = New VbsJson
sRunStopProp = objFSO.OpenTextFile(sPluginPath & "runstop.prop").ReadAll
Set objProp = cJson.Decode(sRunStopProp)
bUseKinect             = CBool(objProp("modules")("runstop")("UseKinect"))
bUseKinectAudio        = CBool(objProp("modules")("runstop")("UseKinectAudio"))
bMinExe                = CBool(objProp("modules")("runstop")("MinimServer"))
bHideExe               = CBool(objProp("modules")("runstop")("HideServer"))
bRunConsole            = CBool(objProp("modules")("runstop")("RunLog2Console"))
bMinConsole            = CBool(objProp("modules")("runstop")("MinimLog2Console"))
iTimeBeforeFirstAction = CInt(objProp("modules")("runstop")("TimeBeforeFirstAction"))
iTimeBetweenEachAction = CInt(objProp("modules")("runstop")("TimeBetweenEachAction"))


if bUseKinect  = true then
	if bUseKinectAudio = true then
		sClient = sKinectAudio
	else
		sClient = sKinect
	end if
	sCmdLineClient=sCmdLineKinect
else
	sClient = sMicro
	sCmdLineClient=sCmdLineMicro
end if

' Get the Actions to run
iAction = 1
Do	
	sActionNum = RightPad(CStr(iAction), 2, "0")
	'sCurrAction = ReadIni(sScriptPath & "Config_RunStop.ini", "RUN_ACTIONS", "RunAction_" & sActionNum) 	
	sCurrAction = objProp("modules")("runstop")("RunAction_" & sActionNum) 	
	sActions = sActions & sCurrAction & ";"
	iAction = iAction + 1
Loop Until sCurrAction=""
sActions = Trim(Left(sActions, Len(sActions)-2))


'-- Run the Log2Console application (or not)

if bRunConsole = true then
	
	' Browse process to know if Log2Console is already running
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
		Return = WshShell.Run(sConsolePath & sLog2Console, iWindowState, False)
	end if
	
end if


'-- Browse process to know if Server and Client are already running

Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
For Each objProcess in colProcesses
	if not IsNull(objProcess.CommandLine) then		
		' Process for Server NodeJS
		if InStr(1, objProcess.CommandLine, sCmdLineNode) <> 0 then
			bNodeIsRunning = true
		end if
		' Process for Client (Kinect or Microphone)
		if InStr(1, objProcess.CommandLine, sCmdLineClient) <> 0 then
			bClientIsRunning = true
		end if
	end if
Next 
Set colProcesses = nothing


'-- Launch executables (or not) with their window state = Minimized/Hidden (or not)

WshShell.CurrentDirectory = sRootPath
if bMinExe = true then
	iWindowState = 7
else
	iWindowState = 1
end if
if bHideExe = true then
	iWindowState = 0
end if
	
if bNodeIsRunning = false then
	' Launch Server
	Return = WshShell.Run(sRootPath & sNodeJS, iWindowState, False)
end if

if bClientIsRunning = false then
	' Refresh the SystemTray icons if the Micro/Kinect has been previously killed, to remove orphaned (ghost) icons
	Return = WshShell.Run(sScriptPath & sAutoItRefresh, 1, true)
	' Launch Client
	Return = WshShell.Run(sRootPath & sClient, iWindowState, False)
end if


'-- Run the actions : Launch some executables, some plugins or say vocal messages

if sActions <> "" then
	
	' Wait for X ms before speaking the welcome sentence
	WScript.sleep iTimeBeforeFirstAction	
	
	aActions = Split(sActions, ";")	
	for each sActionURL in aActions 
	
		if UCase(Left(sActionURL, 4)) = "SAY:" then
			bIsSentence = true
			bIsPlugin = false
			bIsExe = false
			bIsEmulation = false
		elseif UCase(Left(sActionURL, 4)) = "PLG:" then
			bIsSentence = false
			bIsPlugin = true
			bIsExe = false
			bIsEmulation = false
		elseif UCase(Left(sActionURL, 4)) = "EXE:" then
			bIsSentence = false
			bIsPlugin = false
			bIsExe = true
			bIsEmulation = false
		elseif UCase(Left(sActionURL, 4)) = "EMU:" then
			bIsSentence = false
			bIsPlugin = false
			bIsExe = false
			bIsEmulation = true
		else
			bIsSentence = false
			bIsPlugin = false
			bIsExe = false
			bIsEmulation = false
		end if
		
		Return = ""
		
		if bIsExe = true then
			' Launch the executable, or web site
			sActionURL = Right(sActionURL, Len(sActionURL) - 4)
			iReturnValue = WshShell.Run(sActionURL, 1, false)			
		else
			if (bIsSentence = true) or (bIsPlugin = true) or (bIsEmulation = true) then
				sActionURL = Right(sActionURL, Len(sActionURL) - 4)
			
			 ' Wait for 100ms that the Logs have been correctly written
				'WScript.sleep 100
				
				' Wait until SARAH is NOT speaking before launching the current Plugin
				WshShell.CurrentDirectory = sScriptPath
				IsSpeaking = WshShell.Run(sScriptPath & sIsSpeakingLoop & " /timeloop:" & CStr(DEF_TIMELOOP) & " /timeout:" & CStr(DEF_TIMEOUT), 1, true)
				
				' Time to wait between each Action
				WScript.sleep iTimeBetweenEachAction
			end if
	
			if (bIsPlugin = true) then
				WshShell.CurrentDirectory = sRootPath
				' http://127.0.0.1:8080/sarah/pluginname?action=value
				sURL = "http://" & _
				       ReadIni(sRootPath & "custom.ini", "nodejs", "server") & ":" & _
				       ReadIni(sRootPath & "custom.ini", "nodejs", "port") & "/sarah/"
				sToSend = sURL & sActionURL
				Return = SendHTTP(sToSend)
			else
				Return = sActionURL
			end if
	
			if (bIsPlugin = true) or (bIsSentence = true) then
				' http://127.0.0.1:8888/?tts=XXXXXXXXXX
				sURL = "http://" & _
				       ReadIni(sRootPath & "custom.ini", "nodejs", "server") & ":" & _
				       ReadIni(sRootPath & "custom.ini", "common", "loopback") & "/?tts="
				sToSend = sURL & Return
				Return = SendHTTP(sToSend)
			end if
			
			if (bIsEmulation = true) then
				' http://127.0.0.1:8888/?emulate=XXXXXXXXXX
				sURL = "http://" & _
				       ReadIni(sRootPath & "custom.ini", "nodejs", "server") & ":" & _
				       ReadIni(sRootPath & "custom.ini", "common", "loopback") & "/?emulate="
				sToSend = sURL & sActionURL
				Return = SendHTTP(sToSend)
			end if
			
		end if ' bIsExe = true
		
	next
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
' LEFT PADDING
'--------------------------------------

Function LeftPad(sText, iLen, chrPad)
    'LeftPad( "1234", 7, "x" ) = "1234xxx"
    'LeftPad( "1234", 3, "x" ) = "123"
    LeftPad = Left(sText & String(iLen, chrPad), iLen)
End Function


'--------------------------------------
' RIGHT PADDING
'--------------------------------------

Function RightPad(sText, iLen, chrPad )
    'RightPad( "1234", 7, "x" ) = "xxx1234"
    'RightPad( "1234", 3, "x" ) = "234"
    RightPad = Right(String(iLen, chrPad) & sText, iLen)
End Function


'--------------------------------------
' INCLUDE OTHER VBS LIBRARIES
'--------------------------------------

Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub