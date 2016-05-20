#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         BillNyeTheScienceGuy

 Script Function:
	This script switches the active window to Program Manager whenever the mouse 
	hovers over the task bar and when the active program is not "Program 
	Manager," "Start menu," "Jump List," and "".  This allows other scripts that 
	interrupt taskbar options repeatedly to no longer do this.  The only 
	downside is that mousing over the taskbar makes the current active window 
	not the active window anymore.  This can be counteracted by clicking the 
	window again.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

; Gets the height of the taskbar
Opt("WinTitleMatchMode", 4)
Local $iPos = WinGetPos("classname=Shell_TrayWnd")
Local $iTaskbarPixelHeight = $iPos[3]

Local $iTopOfTaskbar = @DesktopHeight - $iTaskbarPixelHeight ; pixel value of top border of taskbar
Local $iSideOfTaskbar = @DesktopWidth ; pixel value of right border of taskbar

While True
	If MouseGetPos(0) <= $iSideOfTaskbar And MouseGetPos(0) >= 54 And MouseGetPos(1) >= $iTopOfTaskbar Then
		; Checking for "Program Manager" and "Start menu" allows start menu to open uniterrupted (like normal operation)
		; MsgBox(0, "", WinGetTitle("[active]"))
		If WinGetTitle("[active]") <> "Program Manager" And WinGetTitle("[active]") <> "Start menu" And WinGetTitle("[active]") <> "Jump List" And WinGetTitle("[active]") <> "" Then
			WinActivate("Program Manager") ; activates taskbar (desktop), allowing taskbar thumbnail previews to work
		EndIf
	EndIf
	
	Sleep(100) ; reduce CPU usage
WEnd
