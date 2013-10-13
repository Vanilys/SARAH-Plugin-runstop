'=====================================
' LOOP TEST UNTIL SARAH STOPS SPEAKING
'=====================================
' ARGUMENTS :
' Format : /name:value
' Available arguments : 
' 	timeloop (ms) : Time to wait between each test inside the loop. By default : 100 ms
' 	timeout (ms) : maximim time for tests. By default : 30000 ms
' So it looks like :
' IsSpeaking_LoopTest.vbs /timeloop:100 /timeout:30000

' RETURN VALUES :
' -1 => LogFile doesn't exist or there is an error
'  0 => SARAH is NOT speaking
'  1 => Time Out occured : max time has been reached

'--------------------------------------
' Main procedure
'--------------------------------------

Dim colNamedArguments, WshShell
Dim iReturnValue, iTimeloop, iTimeOut, iIsSpeaking
Dim sScriptPath
Dim TimeBegin

Const DEFAULT_TIMELOOP = 100
Const DEFAULT_TIMEOUT = 30000

'-- Initialize parameters

iReturnValue = -1
sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
sIsSpeaking  = "IsSpeaking.vbs"

'-- Read the arguments

Set colNamedArguments = WScript.Arguments.Named

If colNamedArguments.Exists("timeloop") Then
	iTimeloop = CInt(colNamedArguments.Item("timeloop"))
Else
	iTimeloop = DEFAULT_TIMELOOP
End If
If colNamedArguments.Exists("timeout") Then
	iTimeOut = CInt(colNamedArguments.Item("timeout"))
Else
	iTimeOut = DEFAULT_TIMEOUT
End If
iTimeOut = iTimeOut / 1000

Set colNamedArguments = nothing


'-- Check that SARAH is speaking or not

Set WshShell = WScript.CreateObject("WScript.Shell")

TimeBegin = Time
Do
	' Wait for 100 ms between each test
	WScript.Sleep(iTimeloop)

	'	ISSPEAKING RETURN VALUES :
	'	 -1 => LogFile doesn't exist or there is an error
	'	  0 => SARAH just finished to speak
	'		1 => SARAH is speaking
	'		2 => SARAH is doing something else
	iIsSpeaking = WshShell.Run(sScriptPath & sIsSpeaking, 1, true)
	
	Select case iIsSpeaking
		Case -1
			iReturnValue = -1
		Case 0
			iReturnValue = 0
		Case 2
			iReturnValue = 0
	End Select
	
	if DateDiff("s", TimeBegin, Time) > iTimeOut then
		' Exit speaking test if it takes more than 30 sec
		iReturnValue = 1
		Exit Do
	end if

Loop Until (iIsSpeaking <> 1)
		
		
'-- Destroy objects
Set WshShell = nothing

WScript.Quit(iReturnValue)
