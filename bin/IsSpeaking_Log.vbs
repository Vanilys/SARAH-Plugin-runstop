'====================================
' SCRIPT TO KNOW IF SARAH IS SPEAKING
' BASED ON LOG FILES
'====================================

' RETURN VALUES :
' -1 => LogFile doesn't exist or there is an error
'  0 => SARAH just finished to speak
'  1 => SARAH is speaking
'  2 => SARAH is doing something else


'--------------------------------------
' Main procedure
'--------------------------------------

Dim sScriptPath, sRootPath, sLogPath
Dim WshShell, objFSO, objTextFile
Dim arrFileLines()
Dim i, iReturnValue
Dim sLastLine, sFileName
Dim CurrTime, CurrDate


'-- Initialize parameters

iReturnValue = -1
sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sRootPath    = Replace(sScriptPath, "plugins\runstop\bin\", "")
sLogPath = sRootPath & "bin\"

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

CurrDate = Date
sFileName = CStr(Year(CurrDate)) & "-" & RightPad(CStr(Month(CurrDate)), 2, "0") & "-" & RightPad(CStr(Day(CurrDate)), 2, "0") & ".log"  
WshShell.CurrentDirectory = sLogPath


'-- Read the Log File

if objFSO.FileExists(sFileName) then

	' Open the Text file. Attributs: read only mode (1) and no file creation if it doesn't exist (False)
	Set objTextFile = objFSO.OpenTextFile(sFileName, 1, False)
	
	' Store all the lines in an array
	Do Until objTextFile.AtEndOfStream 
		Redim Preserve arrFileLines(i)
		arrFileLines(i) = objTextFile.ReadLine
		i = i + 1
	Loop

	objTextFile.Close
	Set objTextFile = nothing

	' Get the last line
	sLastLine = arrFileLines(Ubound(arrFileLines))

	' Return value : 1 if the last line looks like "[00:55:21] [PLAYER]	 [0]Start Player"
	'                0 if the last line looks like "[00:55:40] [PLAYER]	 [0]End Player"
	'                2 if the last line looks like anything else
	if Mid(sLastLine, 12, 8) = "[PLAYER]" and _
	   Right(sLastLine, Len("Start Player")) = "Start Player" then
		iReturnValue = 1
	else 
		if Mid(sLastLine, 12, 8) = "[PLAYER]" and _
		   Right(sLastLine, Len("End Player")) = "End Player" then
			iReturnValue = 0
		else 
			iReturnValue = 2
		end if			
	end if

end if


'-- Destroy objects

Set objFSO = nothing
Set WshShell = nothing

WScript.Quit(iReturnValue)



'--------------------------------------
' RIGHT PADDING
'--------------------------------------

Function RightPad(sText, iLen, chrPad )
    'RightPad( "1234", 7, "x" ) = "xxx1234"
    'RightPad( "1234", 3, "x" ) = "234"
    RightPad = Right(String(iLen, chrPad) & sText, iLen)
End Function