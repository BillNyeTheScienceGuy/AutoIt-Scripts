#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         BillNyeTheScienceGuy

 Script Function:
	This script switches the active window to Program Manager whenever the mouse 
	hovers over the task bar and when the active program is not "Program 
	Manager," "Start menu," "Jump List," and "".  This allows other scripts that 
	interrupt taskbar options repeatedly to no longer do this.  One	downside is 
	that mousing over the taskbar makes the current active window not the active 
	window anymore.  This can be counteracted by clicking the window again.  
	Another downside is that windows can't be minimized by clicking the taskbar 
	icon.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

; Gets the height of the taskbar
Opt("WinTitleMatchMode", 4)
Local $iPos = WinGetPos("classname=Shell_TrayWnd")
Local $iTaskbarPixelHeight = $iPos[3]
Local $iTopOfTaskbar = @DesktopHeight - $iTaskbarPixelHeight ; pixel value of top border of taskbar
Local $iSideOfTaskbar = @DesktopWidth ; pixel value of right border of taskbar
Local $sActiveWindowTitle

While True
	If MouseGetPos(0) <= $iSideOfTaskbar And MouseGetPos(0) >= 54 And MouseGetPos(1) >= $iTopOfTaskbar Then
		; Checking for "Program Manager" and "Start menu" allows start menu to open uniterrupted (like normal operation)
		; MsgBox(0, "", WinGetTitle("[active]"))
		$sActiveWindowTitle = WinGetTitle("[active]")
		If $sActiveWindowTitle <> "Program Manager" And $sActiveWindowTitle <> "Start menu" And $sActiveWindowTitle <> "Jump List" And $sActiveWindowTitle <> "" And $sActiveWindowTitle <> "Date and Time Information" Then
			WinActivate("Program Manager") ; activates taskbar (desktop), allowing taskbar thumbnail previews to work
		EndIf
	EndIf
	
	Sleep(100) ; reduce CPU usage
WEnd
