#Include <GuiToolBar.au3>
#NoTrayIcon

If StringInStr($cmdlineraw, '/Speech') Then
	Global $sToolTipTitle = "SARAH Speech Recognition" ;
ElseIf StringInStr($cmdlineraw, '/Console') Then
	Global $sToolTipTitle = "Log2Console" ;
Else
	Global $sToolTipTitle = ""
EndIf

Global $hSysTray_Handle, $iSystray_ButtonNumber


;~ WinActivate('[Class:Shell_TrayWnd]') ; if taskbar is Autohide

;Opt('WINTITLEMATCHMODE', 4) 
;ControlHide('classname=Shell_TrayWnd', '', '') 
;ControlShow('classname=Shell_TrayWnd', '', '')



$iSystray_ButtonNumber = Get_Systray_Index($sToolTipTitle)

If $iSystray_ButtonNumber = -1 Then
    ;MsgBox(16, "Error", "Icon not found in system tray")
    Exit (-1)
Else
    Sleep(500)
    _GUICtrlToolbar_ClickButton($hSysTray_Handle, $iSystray_ButtonNumber, "right")
    
    ; Browse the menu and click on the first bottom item
    Sleep(10)
    Send("{UP}{ENTER}")
EndIf


Exit (0)


Func Get_Systray_Index($sToolTipTitle)

    ; Find systray handle
    $hSysTray_Handle = ControlGetHandle('[Class:Shell_TrayWnd]', '', '[Class:ToolbarWindow32;Instance:1]')
    If @error Then
        ;MsgBox(16, "Error", "System tray not found")
        Exit (-1)
    EndIf

    ; Get systray item count
    Local $iSystray_ButCount = _GUICtrlToolbar_ButtonCount($hSysTray_Handle)
    If $iSystray_ButCount = 0 Then
        ;MsgBox(16, "Error", "No items found in system tray")
        Exit (-1)
    EndIf

    ; Look for wanted tooltip
    For $iSystray_ButtonNumber = 0 To $iSystray_ButCount
        If StringInStr(_GUICtrlToolbar_GetButtonText($hSysTray_Handle, $iSystray_ButtonNumber), $sToolTipTitle) <> 0 Then 
            ;MsgBox(0,$iSystray_ButCount,$iSystray_ButtonNumber)
            Return $iSystray_ButtonNumber ; Found
        EndIf
    Next
    Return -1 ; Not found

EndFunc