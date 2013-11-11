#NoTrayIcon
#include <Lib_SysTray.au3>
#include <Process.au3>

$count = _SysTrayIconCount()
ConsoleWrite("Count visible tray:  " & $count & @CRLF)
For $i = $count - 1 To 0 Step -1
    $handle = _SysTrayIconHandle($i)
    $visible = _SysTrayIconVisible($i)
    $pid = WinGetProcess($handle)
    $name = _ProcessGetName($pid)
    $title = WinGetTitle($handle)
    $tooltip = _SysTrayIconTooltip($i)
    If $pid = -1 Then _SysTrayIconRemove($i)
Next

If _FindTrayToolbarWindow(2) <> -1 Then
    $countwin7 = _SysTrayIconCount(2)
    For $i = $countwin7 - 1 To 0 Step -1
        $handle = _SysTrayIconHandle($i, 2)
        $visible = _SysTrayIconVisible($i, 2)
        $pid = WinGetProcess($handle)
        $name = _ProcessGetName($pid)
        $title = WinGetTitle($handle)
        $tooltip = _SysTrayIconTooltip($i, 2)
        If $pid = -1 Then _SysTrayIconRemove($i, 2)
    Next
EndIf