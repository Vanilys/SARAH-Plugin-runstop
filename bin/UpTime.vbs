'===================================
' GET LAST BOOT AND UPTIME
' OF SYSTEM AND SARAH
'===================================

' RETURN VALUES (JSon file) :
' Last boot system
' Sytem uptime
' Last boot SARAH
' SARAH uptime

' The results are like :
'{
'	"infos":{
'		"runstop":{
'			"LastBootSystem":"19/10/2013 23:13:49",
'			"UptimeSystem":"25:31:19",
'			"LastBootSarah":"21/10/2013 19:20:03",
'			"UptimeSarah":"00:25:05"
'			}
'	}
'}


'--------------------------------------
' Main procedure
'--------------------------------------

Dim objFSO, objWMIService, objWMIDateTime, colOS, colProcesses, objProcess
Dim sComputer
Dim sScriptPath, sPluginPath, sRootPath
Dim iReturnValue
Dim sLastBootSystem, sUptimeSystem, sLastBootSarah, sUptimeSarah

iReturnValue = -1
sComputer = "." ' Local computer
sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sPluginPath = Replace(sScriptPath, "bin\", "")
sRootPath    = Replace(sScriptPath, "plugins\runstop\bin\", "")
sSystemProp  = "uptime.json"

includeFile sScriptPath & "Lib_VbsJson.vbs"


Set objFSO = CreateObject( "Scripting.FileSystemObject" )
set objWMIDateTime = CreateObject("WbemScripting.SWbemDateTime")
set objWMIService = GetObject("winmgmts:\\" & sComputer & "\root\cimv2")
set colOS = objWMIService.InstancesOf("Win32_OperatingSystem")


' -- Get Last boot and Uptime for the system

For each objOS in colOS
	objWMIDateTime.Value = objOS.LastBootUpTime
	'Wscript.Echo "Last boot system : " & objWMIDateTime.GetVarDate & vbcrlf & _
	'	"Sytem uptime : " &  TimeSpan(objWMIDateTime.GetVarDate,Now) & _
	'	" (hh:mm:ss)"
	iReturnValue = 0
	sLastBootSystem = objWMIDateTime.GetVarDate
	sUptimeSystem = TimeSpan(objWMIDateTime.GetVarDate,Now)
Next


' -- Get Last boot and Uptime for SARAH

Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")

If objFSO.FileExists(sRootPath & "WSRNode.cmd") Then	
	' SARAH Version <= 3 alpha 2 (including 2.8, 2.95 ...)
	sCmdLineNode = "WSRNode.cmd"	
ElseIf objFSO.FileExists(sRootPath & "Server_NodeJS.cmd") Then	
	' SARAH Version >= 3 beta 1 (including 3RC1, 3RC2, 3.0 ...)
	sCmdLineNode = "Server_NodeJS.cmd"	
Else
	WScript.Echo "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If


' Browse process
For Each objProcess in colProcesses
	if not IsNull(objProcess.CommandLine) then	
		if InStr(1, objProcess.CommandLine, sCmdLineNode) <> 0 then
			objWMIDateTime.Value = objProcess.CreationDate
			'Wscript.Echo "Last boot SARAH : " & objWMIDateTime.GetVarDate & vbcrlf & _
			'	"SARAH uptime : " &  TimeSpan(objWMIDateTime.GetVarDate,Now) & _
			'	" (hh:mm:ss)"
			iReturnValue = 0
			sLastBootSarah = objWMIDateTime.GetVarDate
			sUptimeSarah = TimeSpan(objWMIDateTime.GetVarDate,Now)
		end if
	end if
Next 


' -- Write the results in the JSON file

Dim jProp
Set jProp = CreateObject("Scripting.Dictionary")
jProp.Add "LastBootSystem", sLastBootSystem
jProp.Add "UptimeSystem", sUptimeSystem
jProp.Add "LastBootSarah", sLastBootSarah
jProp.Add "UptimeSarah", sUptimeSarah

Dim jRunstop
Set jRunstop = CreateObject("Scripting.Dictionary")
jRunstop.Add "runstop", jProp

Dim jRoot
Set jRoot = CreateObject("Scripting.Dictionary")
jRoot.Add "infos", jRunstop

Dim cJson, objProp
Dim sResult
Set cJson = New VbsJson
sResult = cJson.Encode(jRoot)
Set objProp = objFSO.CreateTextFile(sPluginPath & sSystemProp, True, False)
objProp.WriteLine sResult
objProp.Close



'-- Destroy objects 

Set objProp = nothing
Set cJson = nothing
set jProp = nothing
set jRunstop = nothing
set jRoot = nothing

Set colProcesses = nothing
Set objProcess = nothing
set colOS = nothing
set objWMIService = nothing
set objWMIDateTime = nothing


WScript.Quit(iReturnValue)



'--------------------------------------
' Display the difference between 2 dates
' in hh:mm:ss format
'--------------------------------------

Function TimeSpan(dt1, dt2) 

	If (isDate(dt1) And IsDate(dt2)) = false Then 
		TimeSpan = "00:00:00" 
		Exit Function 
	End If 
 
	seconds = Abs(DateDiff("S", dt1, dt2)) 
	minutes = seconds \ 60 
	hours = minutes \ 60 
	minutes = minutes mod 60 
	seconds = seconds mod 60 
	
	if len(hours) = 1 then hours = "0" & hours 
	
	TimeSpan = hours & ":" & _ 
	    RIGHT("00" & minutes, 2) & ":" & _ 
	    RIGHT("00" & seconds, 2) 
End Function 


'--------------------------------------
' INCLUDE OTHER VBS LIBRARIES
'--------------------------------------

Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub