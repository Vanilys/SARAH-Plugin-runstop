'==============================
' STOP S.A.R.A.H. SERVER
'==============================
Option explicit 

'--------------------------------------
' Main procedure
'--------------------------------------

Dim WshShell, objWMIService, objFSO, colProcesses, objProcess, objPro
Dim sScriptPath, sPluginPath, sRootPath, sCmdLineNode
Dim bStopServer
Dim iReturnValue, NodePID


'-- Initialize parameters

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set objFSO = CreateObject( "Scripting.FileSystemObject" )

iReturnValue = -1
sScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sPluginPath = Replace(sScriptPath, "bin\", "")
sRootPath   = Replace(sScriptPath, "plugins\runstop\bin\", "")


If objFSO.FileExists(sRootPath & "WSRNode.cmd") Then
	
	' SARAH Version <= 3 alpha 2 (including 2.8, 2.95 ...)
	sCmdLineNode        = "WSRNode.cmd"
	
ElseIf objFSO.FileExists(sRootPath & "Server_NodeJS.cmd") Then
	
	' SARAH Version >= 3 beta 1 (including 3RC1, 3RC2, 3.0 ...)
	sCmdLineNode        = "Server_NodeJS.cmd"
	
Else
	Wscript.Echo "Impossible de détecter la version de SARAH."
	WScript.Quit(iReturnValue)
End If


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
				
	end if
Next


'--  Destroy objects

Set colProcesses = nothing
Set objWMIService = nothing
Set WshShell = nothing
Set objFSO = nothing


WScript.Quit(iReturnValue)



'--------------------------------------
' INCLUDE OTHER VBS LIBRARIES
'--------------------------------------

Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub