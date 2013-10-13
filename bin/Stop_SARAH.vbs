'==============================
' SCRIPT TO STOP S.A.R.A.H
'==============================

'--------------------------------------
' Main procedure
'--------------------------------------

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess, objPro
Dim sScriptPath, sRootPath, sCmdLineNode, sCmdLineKinect, sCmdLineMacro, sCmdLineConhost, sCmdLineLog2Console, sAutoItExit, sAutoItRefresh
Dim sExitTrayIcon_Speech, sExitTrayIcon_Console
Dim bStopConsole, bStopGracefully
Dim iReturnValue, iTimeBeforeKill
Dim TimeBegin


'-- Initialize parameters

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set objFSO = CreateObject( "Scripting.FileSystemObject" )

iReturnValue = -1
sScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sRootPath    = Replace(sScriptPath, "plugins\runstop\bin\", "")
sAutoItExit           = "Exit_TrayIcon.exe"
sAutoItRefresh        = "Refresh_SystemTray.exe"
sExitTrayIcon_Speech  = sAutoItExit & " /Speech"
sExitTrayIcon_Console = sAutoItExit & " /Console"
sCmdLineLog2Console   = "Log2Console.exe"

If objFSO.FileExists(sRootPath & "WSRNode.cmd") Then
	
	' SARAH Version <= 3 alpha 2 (including 2.8, 2.95 ...)
	sCmdLineNode        = "WSRNode.cmd"
	sCmdLineMicro       = "WSRMacro.exe"
	sCmdLineKinect      = "WSRMacro_Kinect.exe"
	
ElseIf objFSO.FileExists(sRootPath & "Server_NodeJS.cmd") Then
	
	' SARAH Version >= 3 beta 1
	sCmdLineNode        = "Server_NodeJS.cmd"
	sCmdLineMicro       = "WSRMacro.exe"
	sCmdLineKinect      = "WSRMacro_Kinect.exe"
	
Else
	MsgBox "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If


'-- Read ini file values

bStopConsole        = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "STOP", "StopLog2Console"))
bStopGracefully     = CBool(ReadIni(sScriptPath & "Config_RunStop.ini", "STOP", "StopGracefully"))
iTimeBeforeKill     = CInt(ReadIni(sScriptPath & "Config_RunStop.ini", "STOP", "TimeBeforeKill"))
if iTimeBeforeKill > 1000 then iTimeBeforeKill = 60


'-- Browse all of the process, and check their command line

WshShell.CurrentDirectory = sScriptPath
Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
For Each objProcess in colProcesses
	if not IsNull(objProcess.CommandLine) then
	
		' Process for NodeJS
		if InStr(1, objProcess.CommandLine, sCmdLineNode) <> 0 then
			NodePID = objProcess.ProcessId
			'msgbox objProcess.ProcessId & " | " & objProcess.ParentProcessId & " | " & objProcess.CommandLine
			objProcess.Terminate()
			
		' Child Process of NodeJS 
			For Each objPro in colProcesses
				if objPro.ParentProcessId = NodePID then
					'msgbox objPro.ProcessId & " | " & objPro.ParentProcessId & " | " & objPro.CommandLine
					objPro.Terminate()
				end if
			Next				
		end if
		
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
		
		' Process for the console
		if bStopConsole = true then
			if InStr(1, objProcess.CommandLine, sCmdLineLog2Console) <> 0 then
				'msgbox objProcess.ProcessId & " | " & objProcess.ParentProcessId & " | " & objProcess.CommandLine	
				if bStopGracefully = false then 
					objProcess.Terminate()
				else
					Return = WshShell.Run(sScriptPath & sExitTrayIcon_Console, 1, true)
				end if
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
				
				' Console
				if bStopConsole = true then
					if InStr(1, objProcess.CommandLine, sCmdLineLog2Console) <> 0 then
						iNbProcessRunning = iNbProcessRunning + 1
						if DateDiff("s", TimeBegin, Time) >= iTimeBeforeKill then
							objProcess.Terminate()
						end if
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

Set colProcesses = nothing
Set objWMIService = nothing
Set WshShell = nothing
Set objFSO = nothing


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