'==============================
' SCRIPT TO RUN S.A.R.A.H
'==============================

'--------------------------------------
' Main procedure
'--------------------------------------
Dim sCurrentPath, sScriptPath, sConsolePath, sNodeJS, sKinect, sMicro, sLog2Console, sSpeaker
Dim iWindowState
Dim bConsoleIsRunning
Dim sURL, sMessage, sToSend
Dim iReturnValue


iReturnValue = -1

sScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sRootPath = Replace(sScriptPath, "\plugins\runstop\bin", "")
sConsolePath = sRootPath & "\Log2Console"

' Read in file values
sNodeJS = ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "NodeJS")
sMicro = ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "Micro")
sKinect = ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "Kinect")
sLog2Console = ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "Log2Console")

if ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "UseKinect") = "true" then
	sSpeaker = sKinect
else
	sSpeaker = sMicro
end if


Set WshShell = WScript.CreateObject("WScript.Shell")

' Run the Log2Console application (or not)
if ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "RunLog2Console") = "true" then
	
	' Browse process to know if Log2Console is already running
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
	Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
	For Each objProcess in colProcesses
		if not IsNull(objProcess.CommandLine) then		
			' Process for Log2Console
			if InStr(1, objProcess.CommandLine, sLog2Console) <> 0 then
				bConsoleIsRunning = true
			end if
		end if
	Next 
	Set colProcesses = nothing
	Set objWMIService = nothing

	' If Log2Console isn't already running, then run it !
	if bConsoleIsRunning = false then
		WshShell.CurrentDirectory = sConsolePath
		if ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "MinLog2Console") = "true" then
			iWindowState = 7
		else
			iWindowState = 1
		end if
		Return = WshShell.Run(sConsolePath & "\" & sLog2Console, iWindowState, False)
	end if
	
end if

' Launch executables with their window state = Minimized or not
WshShell.CurrentDirectory = sRootPath
if ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "MinExe") = "true" then
	iWindowState = 7
else
	iWindowState = 1
end if
Return = WshShell.Run(sRootPath & "\" & sNodeJS, iWindowState, False)
Return = WshShell.Run(sRootPath & "\" & sSpeaker, iWindowState, False)

Set WshShell = nothing


' Send the sentence to SARAH
sMessage = ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "Speech")

if sMessage <> "" then
	' Wait for X ms before speaking the welcome sentence
	WScript.sleep ReadIni(sScriptPath & "\Config_RunStop.ini", "RUN", "WaitBeforeSpeech")	
	
	' http://127.0.0.1:8888/?tts=XXXXXXXXXX
	sURL = "http://" & _
	       ReadIni(sRootPath & "\custom.ini", "nodejs", "server") & ":" & _
	       ReadIni(sRootPath & "\custom.ini", "common", "loopback") & "/?tts="
	sToSend = sURL & sMessage
	
	Set xmlHttp = WScript.CreateObject("MSXML2.ServerXMLHTTP")
	xmlHttp.Open "GET", sToSend, False
	xmlHttp.Send ""
	getHTML = xmlHttp.responseText
	status = xmlHttp.status
	xmlHttp.Abort
	Set xmlHttp = Nothing
	
	If status = 200 Then
		If Len(getHTML) > 0 Then
			'...
		End If
	End If
	
end if


iReturnValue = 0
WScript.Quit(iReturnValue)


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
End Function