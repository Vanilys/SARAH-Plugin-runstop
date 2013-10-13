'==============================
' SCRIPT TO RUN S.A.R.A.H
'==============================

'--------------------------------------
' Main procedure
'--------------------------------------

Dim sScriptPath, sRootPath, sConsolePath, sIsSpeakingLoop
Dim sNodeJS, sKinect, sMicro, sLog2Console, sClient
Dim sCmdLineNode, sCmdLineKinect, sCmdLineMicro, sCmdLineConhost, sCmdLineLog2Console, sCmdLineClient
Dim iWindowState, iTimeBeforeFirstAction, iTimeBetweenEachAction
Dim bUseKinect, bUseKinectAudio, bMinExe, bHideExe, bRunConsole, bMinConsole, bConsoleIsRunning, bNodeIsRunning, bClientIsRunning
Dim sURL, sToSend
Dim iReturnValue, IsSpeaking, iAction
Dim sActions, sCurrAction, sActionNum, sActionURL
Dim aActions
Dim bIsPlugin, bIsSentence, bIsExe
Dim TimeBegin

Const DEF_TIMELOOP = 100
Const DEF_TIMEOUT = 30000

'-- Initialize parameters
		
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set objFSO = CreateObject( "Scripting.FileSystemObject" )

iReturnValue = -1
sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sRootPath    = Replace(sScriptPath, "plugins\runstop\bin\", "")
sConsolePath = sRootPath & "Log2Console\"
sIsSpeakingLoop     = "IsSpeaking_LoopTest.vbs"
sAutoItRefresh      = "Refresh_SystemTray.exe"
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
	
	' SARAH Version >= 3 beta 1
	sNodeJS           = "Server_NodeJS.cmd"
	sMicro            = "Client_Microphone.cmd"
	sKinect           = "Client_Kinect.cmd"
	sKinectAudio      = "Client_Kinect_Audio.cmd"
	sCmdLineNode      = "Server_NodeJS.cmd"
	sCmdLineMicro     = "WSRMacro.exe"
	sCmdLineKinect    = "WSRMacro_Kinect.exe"
	
Else
	MsgBNox "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If


'-- Read ini file values

bUseKinect             = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "UseKinect"))
bUseKinectAudio        = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "UseKinectAudio"))
bMinExe                = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "MinimExe"))
bHideExe               = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "HideExe"))
bRunConsole            = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "RunLog2Console"))
bMinConsole            = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN", "MinimLog2Console"))
iTimeBeforeFirstAction = CInt(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN_ACTIONS", "TimeBeforeFirstAction"))
iTimeBetweenEachAction = CInt(ReadIni(sScriptPath & "Config_RunStop.ini", "RUN_ACTIONS", "TimeBetweenEachAction"))

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
	sCurrAction = ReadIni(sScriptPath & "Config_RunStop.ini", "RUN_ACTIONS", "RunAction_" & sActionNum) 	
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


'-- Run the acioons : Launch some executables, some plugins or say vocal messages

if sActions <> "" then
	
	' Wait for X ms before speaking the welcome sentence
	WScript.sleep iTimeBeforeFirstAction	
	
	aActions = Split(sActions, ";")	
	for each sActionURL in aActions 
	
		if UCase(Left(sActionURL, 4)) = "SAY:" then
			bIsSentence = true
			bIsPlugin = false
			bIsExe = false
		elseif UCase(Left(sActionURL, 4)) = "PLG:" then
			bIsSentence = false
			bIsPlugin = true
			bIsExe = false
		elseif UCase(Left(sActionURL, 4)) = "EXE:" then
			bIsSentence = false
			bIsPlugin = false
			bIsExe = true
		else
			bIsSentence = false
			bIsPlugin = false
			bIsExe = false
		end if
		
		if bIsExe = true then
			' Launch the executable, or web site
			sActionURL = Right(sActionURL, Len(sActionURL) - 4)
			iReturnValue = WshShell.Run(sActionURL, 1, false)			
		else
			if (bIsSentence = true) or (bIsPlugin = true) then
				sActionURL = Right(sActionURL, Len(sActionURL) - 4)
			
			 ' Wait for 100ms that the Logs have been correctly written
				WScript.sleep 100
				
				' Wait until SARAH is NOT speaking before launching the current Plugin
				WshShell.CurrentDirectory = sScriptPath
				IsSpeaking = WshShell.Run(sScriptPath & sIsSpeakingLoop & " /timeloop:" & CStr(DEF_TIMELOOP) & " /timeout:" & CStr(DEF_TIMEOUT), 1, true)
				
				' Time to wait between each Action
				WScript.sleep iTimeBetweenEachAction
			end if
	
			if (bIsPlugin = true) and (bIsSentence = false) then
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
	
			if (bIsSentence = true) or (bIsPlugin = true) then
				' http://127.0.0.1:8888/?tts=XXXXXXXXXX
				sURL = "http://" & _
				       ReadIni(sRootPath & "custom.ini", "nodejs", "server") & ":" & _
				       ReadIni(sRootPath & "custom.ini", "common", "loopback") & "/?tts="
				sToSend = sURL & Return
				Return = SendHTTP(sToSend)
			end if
			
		end if ' bIsExe = true
		
	next
end if


'-- Destroy objects

Set objWMIService = nothing
Set WshShell = nothing
Set objFSO = nothing


iReturnValue = 0
WScript.Quit(iReturnValue)



'--------------------------------------
' SEND URL TO HTTP SERVER
'--------------------------------------

Function SendHTTP(sRequest)

	Set xmlHttp = WScript.CreateObject("MSXML2.ServerXMLHTTP")

	xmlHttp.Open "GET", sRequest, False
	xmlHttp.Send ""
	getHTML = xmlHttp.responseText
	status = xmlHttp.status
	xmlHttp.Abort

	Set xmlHttp = Nothing
	
	If status = 200 Then
		If Len(getHTML) > 0 Then
			SendHTTP = getHTML
		else 
			SendHTTP = "Erreur serveur, retour vide"
		End If
	else
			SendHTTP = "Erreur serveur, status " & status
	End If


End Function



'--------------------------------------
' READ THE INI FILE
'--------------------------------------

Function ReadIni( myFilePath, mySection, myKey )
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    If Left( strLine, 1) <> ";" then
                        ' Find position of equal sign in the line
                        intEqualPos = InStr( 1, strLine, "=", 1 )
                        If intEqualPos > 0 Then
                            strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                            ' Check if item is found in the current line
                            If LCase( strLeftString ) = LCase( strKey ) Then
                                ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                                ' In case the item exists but value is blank
                                If ReadIni = "" Then
                                    ReadIni = " "
                                End If
                                ' Abort loop when item is found
                                Exit Do
                            End If
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If

    Set objIniFile = nothing
    Set objFSO = nothing

End Function




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
