'====================================
' SCRIPT TO KNOW IF SARAH IS SPEAKING
' BASED ON STATUS REQUEST
'====================================

' RETURN VALUES :
' -1 => There is an error (Server ...)
'  0 => SARAH  is NOT speaking
'  1 => SARAH is speaking


'--------------------------------------
' Main procedure
'--------------------------------------

Dim sScriptPath, sRootPath
Dim iReturnValue

'-- Initialize parameters

iReturnValue = -1
sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sRootPath    = Replace(sScriptPath, "plugins\runstop\bin\", "")

includeFile sScriptPath & "Lib_VbsReadIni.vbs"
includeFile sScriptPath & "Lib_VbsSendHTTP.vbs"

' http://127.0.0.1:8888/?status=true
sToSend = "http://127.0.0.1:" & _
           ReadIni(sRootPath & "custom.ini", "common", "loopback") & "/?status=true"
Return = Trim(SendHTTP(sToSend))

if UCase(Return) = "SPEAKING" then
	iReturnValue = 1
elseif (Return  = "") or IsEmpty(Return) or IsNull(Return) then
	iReturnValue = 0
else
	iReturnValue = -1
end if


WScript.Quit(iReturnValue)



'--------------------------------------
' INCLUDE OTHER VBS LIBRARIES
'--------------------------------------

Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub