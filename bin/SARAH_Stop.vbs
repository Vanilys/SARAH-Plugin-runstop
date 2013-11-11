'=====================================
' STOP S.A.R.A.H.
'=====================================
Option explicit 

Dim WshShell, objFSO
Dim sScriptPath, sRootPath, sPluginPath, sStopServer, sStopClient, sStopLog2Console
Dim cJson, objProp
Dim sRunStopProp
Dim sCmdLineKinect, sCmdLineMicro, sCmdLineClient
Dim bStopServer, bStopClient, bStopLog2Console
Dim iReturnValue, iTimeBeforeFirstAction, iTimeloop, iTimeOut
Dim TimeBegin
Dim Return

Const DEFAULT_TIMELOOP = 100
Const DEFAULT_TIMEOUT = 30000

On Error Resume Next

'-- Initialize parameters

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject( "Scripting.FileSystemObject" )

iReturnValue = -1
sScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sPluginPath  = Replace(sScriptPath, "bin\", "")
sRootPath    = Replace(sScriptPath, "plugins\runstop\bin\", "")

sStopServer      = "SARAH_StopServer.vbs"
sStopClient      = "SARAH_StopClient.vbs"
sStopLog2Console = "SARAH_StopLog2Console.vbs"

includeFile sScriptPath & "Lib_VbsJson.vbs"


'-- Read JSon file values

Set cJson    = New VbsJson
sRunStopProp = objFSO.OpenTextFile(sRootPath & "custom.prop").ReadAll
Set objProp  = cJson.Decode(sRunStopProp)
bStopServer      = CBool(objProp("modules")("runstop")("StopServer"))
bStopClient      = CBool(objProp("modules")("runstop")("StopClient"))
bStopLog2Console = CBool(objProp("modules")("runstop")("StopLog2Console"))

if (bStopServer = "") or (bStopClient = "") or (bStopLog2Console = "") then
	Set cJson    = New VbsJson
	sRunStopProp = objFSO.OpenTextFile(sPluginPath & "runstop.prop").ReadAll
	Set objProp  = cJson.Decode(sRunStopProp)
	bStopServer      = CBool(objProp("modules")("runstop")("RunServer"))
	bStopClient      = CBool(objProp("modules")("runstop")("RunClient"))
	bStopLog2Console = CBool(objProp("modules")("runstop")("RunLog2Console"))
end if


' -- Stop applications

'-- Stop Client, and wait for the end of the process
if (bStopClient = true) then 
	Return = WshShell.Run(sScriptPath & sStopClient, 1, true)
end if

'-- Stop Server, and wait for the end of the process
if (bStopServer = true) then 
	Return = WshShell.Run(sScriptPath & sStopServer, 1, true)
end if 

'-- Stop Log2Console, and wait for the end of the process
if (bStopLog2Console = true) then 
	Return = WshShell.Run(sScriptPath & sStopLog2Console, 1, true)
end if


'-- Destroy objects

Set objProp = nothing
Set cJson = nothing
Set WshShell = nothing

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