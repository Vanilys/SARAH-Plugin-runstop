'==============================
' RUN ACTIONS FOR S.A.R.A.H. 
'==============================
Option explicit 

'--------------------------------------
' Main procedure
'--------------------------------------

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess
Dim cJson, objProp
Dim sRunStopProp
Dim sScriptPath, sPluginPath, sRootPath, sIsSpeakingLoop
Dim sCmdLineNode, sCmdLineKinect, sCmdLineMicro, sCmdLineClient
Dim iWindowState, iTimeBeforeFirstAction, iTimeBetweenEachAction
Dim bUseKinect, bNodeIsRunning, bClientIsRunning
Dim sURL, sToSend
Dim iReturnValue, IsSpeaking, iAction
Dim sActions, sCurrAction, sActionNum, sActionURL, sURL_Server, sURL_Client, sReturn
Dim aActions
Dim bIsPlugin, bIsSentence, bIsExe, bIsEmulation
Dim TimeBegin

On Error Resume Next

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

sIsSpeakingLoop     = "IsSpeaking_LoopTest.vbs"

If objFSO.FileExists(sRootPath & "WSRNode.cmd") Then
	
	' SARAH Version <= 3 alpha 2 (including 2.8, 2.95 ...)
	sCmdLineNode      = "WSRNode.cmd"
	sCmdLineMicro     = "WSRMacro.exe"
	sCmdLineKinect    = "WSRMacro_Kinect.exe"
	
ElseIf objFSO.FileExists(sRootPath & "Server_NodeJS.cmd") Then
	
	' SARAH Version >= 3 beta 1 (including 3RC1, 3RC2, 3.0 ...)
	sCmdLineNode      = "Server_NodeJS.cmd"
	sCmdLineMicro     = "WSRMacro.exe"
	sCmdLineKinect    = "WSRMacro_Kinect.exe"
	
Else
	Wscript.Echo "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If

'-- Read ini and JSon file values

' http://127.0.0.1:8080/
sURL_Server = "http://" & ReadIni(sRootPath & "custom.ini", "nodejs", "server") & ":" & _
                          ReadIni(sRootPath & "custom.ini", "nodejs", "port")
						  
' http://127.0.0.1:8888
'sURL_Client = "http://" & ReadIni(sRootPath & "custom.ini", "nodejs", "server") & ":" & _
'						   ReadIni(sRootPath & "custom.ini", "common", "loopback")
sURL_Client = "http://127.0.0.1:" & _
                          ReadIni(sRootPath & "custom.ini", "common", "loopback")

Set cJson    = New VbsJson
sRunStopProp = objFSO.OpenTextFile(sRootPath & "custom.prop").ReadAll
Set objProp  = cJson.Decode(sRunStopProp)
bUseKinect             = CBool(objProp("modules")("runstop")("UseKinect"))
iTimeBetweenEachAction = CInt(objProp("modules")("runstop")("TimeBetweenEachAction"))

if (bUseKinect = "") or (iTimeBetweenEachAction = "") then
	Set cJson    = New VbsJson
	sRunStopProp = objFSO.OpenTextFile(sPluginPath & "runstop.prop").ReadAll
	Set objProp  = cJson.Decode(sRunStopProp)
	bUseKinect             = CBool(objProp("modules")("runstop")("UseKinect"))
	iTimeBetweenEachAction = CInt(objProp("modules")("runstop")("TimeBetweenEachAction"))
end if

if bUseKinect  = true then
	sCmdLineClient = sCmdLineKinect
else
	sCmdLineClient = sCmdLineMicro
end if


'-- Get the Actions to run
iAction = 1
Do	
	sActionNum = RightPad(CStr(iAction), 2, "0")
	sCurrAction = objProp("modules")("runstop")("RunAction_" & sActionNum) 
	sActions = sActions & sCurrAction & ";"
	iAction = iAction + 1
Loop Until sCurrAction=""
sActions = Trim(Left(sActions, Len(sActions)-2))


'-- Browse processes to know if Server and Client are already running

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


'-- Run the actions : Launch some executables, some plugins or say vocal messages

if ((bNodeIsRunning = true) and (bClientIsRunning = true)) and (sActions <> "") then
		
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

				' Wait until SARAH is NOT speaking before launching the current Plugin
				WshShell.CurrentDirectory = sScriptPath
				IsSpeaking = WshShell.Run(sScriptPath & sIsSpeakingLoop & " /timeloop:" & CStr(DEF_TIMELOOP) & " /timeout:" & CStr(DEF_TIMEOUT), 1, true)

				' Time to wait between each Action
				WScript.sleep iTimeBetweenEachAction
			end if

			if (bIsPlugin = true) then
				WshShell.CurrentDirectory = sRootPath
				' http://127.0.0.1:8080/sarah/pluginname?action=value
				sURL = sURL_Server & "/sarah/"
				sToSend = sURL & sActionURL
				sReturn = SendHTTP(sToSend)
			else
				sReturn = sActionURL
			end if
	
			if (bIsPlugin = true) or (bIsSentence = true) then
				' http://127.0.0.1:8888/?tts=XXXXXXXXXX
				sURL = sURL_Client & "/?tts="
				sToSend = sURL & sReturn
				Return = SendHTTP(sToSend)
			end if
			
			if (bIsEmulation = true) then
				' http://127.0.0.1:8888/?emulate=XXXXXXXXXX
				sURL = sURL_Client & "/?emulate="
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