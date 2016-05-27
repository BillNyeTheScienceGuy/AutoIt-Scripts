#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         BillNyeTheScienceGuy

 Script Function:
	Displays Outlook calendar appointments on a Dream Cheeky LED message board 
	using the LED Panel Controller program made by Tiago Rodrigues.  Also 
	displays unread messages using a hotkey (ctrl + alt + u).
	
	This script runs better in conjunction with "taskbar activator.au3", which 
	activates the taskbar when mousing over it, allowing for regular use of the 
	taskbar thumbnail preview and taskbar right-clicking.
	
	NOTE: Script must be in the same directory as the LED Panel Controller 
	executable file.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <Date.au3>      ; for _DateDiff, _DateAdd, and _NowCalc
#include <Outlook.au3>   ; for _OutlookOpen and _OutlookGetAppointments
#include <String.au3>    ; for _StringInsert

Func eventIsHappening($sStart, $sEnd)
	; Reverse order so indexes don't get affected by the resulting index shifting
	$sStart = _StringInsert($sStart, ":", 12)
	$sStart = _StringInsert($sStart, ":", 10)
	$sStart = _StringInsert($sStart, " ",  8)
	$sStart = _StringInsert($sStart, "/",  6)
	$sStart = _StringInsert($sStart, "/",  4)
	
	$sEnd = _StringInsert($sEnd, ":", 12)
	$sEnd = _StringInsert($sEnd, ":", 10)
	$sEnd = _StringInsert($sEnd, " ",  8)
	$sEnd = _StringInsert($sEnd, "/",  6)
	$sEnd = _StringInsert($sEnd, "/",  4)
	
	Local $bAtOrAfterStart = False
	Local $bAtOrBeforeEnd = False
	
	; If current time - start time is positive or 0
	If _DateDiff('n', $sStart, _NowCalc()) >= 0 Then
		$bAtOrAfterStart = True
	EndIf
	
	; If end time - current time is positive or 0
	If _DateDiff('n', _NowCalc(), $sEnd) >= 0 Then
		$bAtOrBeforeEnd = True
	EndIf
	
	; Return True iff current time is [start, end]
	If $bAtOrAfterStart And $bAtOrBeforeEnd Then
		Return True
	EndIf
	Return False
EndFunc

Func allEventsHappening(ByRef $aAppointments)
	If IsArray($aAppointments) Then ; if appointments exist
		For $i = 1 To $aAppointments[0][0]
			; For all appointments, replace location with boolean representing if the event is happening
			$aAppointments[$i][3] = eventIsHappening($aAppointments[$i][1], $aAppointments[$i][2])
		Next
	EndIf
EndFunc

Func changesToAppointmentsArrays($aAppointments, $aLastAppointments)
	Local $bChanged = False
	
	If IsArray($aAppointments) And Not IsArray($aLastAppointments) Then
		$bChanged = True
	ElseIf Not IsArray($aAppointments) And IsArray($aLastAppointments) Then
		$bChanged = True
	Else
		If $aAppointments[0][0] == $aLastAppointments[0][0] Then
			For $i = 1 To $aAppointments[0][0]
				For $j = 0 To 7
					If $aAppointments[$i][$j] <> $aLastAppointments[$i][$j] Then
						$bChanged = True
					EndIf
				Next
				; MsgBox(0, "", $aAppointments[$i][0])
			Next
		Else
			$bChanged = True
		EndIf
	EndIf
	
	Return $bChanged
EndFunc

Func startMessage($sMessage)
	; -SetRepeatCount and -DeleteMsgAfterRepeat don't work, so the Sleep function and -DeleteRegex is necessary
	Run("LedDisplayControllerGui.exe -NoServer -NoGui -SetRepeatCount 1 -DeleteMsgAfterRepeat 1 -SendText """ & $sMessage & "  """)
EndFunc

Func stopMessages()
	Run("LedDisplayControllerGui.exe -NoServer -NoGui -DeleteRegex """"")
EndFunc

Func loadAppointments(ByRef $aAppointments)
	Local $oOutlook = _OutlookOpen()
	Local $sDate = @YEAR & "-" & @MON & "-" & @MDAY ; current date in YYYY-MM-DD format
	$aAppointments = _OutlookGetAppointments($oOutlook, "", $sDate & " 00:00", _DateAdd('d', 1, $sDate) & " 00:00", "", 2, "")
EndFunc

Func forceQuitAllProcesses($sProcessName)
	Local $aProcesses = ProcessList($sProcessName)
	For $i = 0 To UBound($aProcesses) - 1
		Run("taskkill /F /IM " & $sProcessName)
	Next
EndFunc

Func restartScript()
    If @Compiled = 1 Then
        Run(FileGetShortName(@ScriptFullPath))
    Else
        Run(FileGetShortName(@AutoItExe) & " " & FileGetShortName(@ScriptFullPath))
    EndIf
	
    Exit
EndFunc

Local $aAppointments, $aLastAppointments
loadAppointments($aAppointments)

HotKeySet("^!r", "restartScript") ; set Ctrl + Alt + R as the hotkey to restart the script

; Quit all previous controller processes, then start LED Controller app server with no GUI
forceQuitAllProcesses("LedDisplayControllerGui.exe")
Sleep(500) ; give process-quiting time to stop so a new process can be created
Run("LedDisplayControllerGui.exe -NoGui") ; start controller server with no GUI

; Loops through the appointments and displays all the ones currently taking place
While True
	
	loadAppointments($aAppointments)
	
	allEventsHappening($aAppointments)
	
	If changesToAppointmentsArrays($aAppointments, $aLastAppointments) Then
		stopMessages()
		
		If IsArray($aAppointments) Then ; if appointments exist
			Sleep(100)
			For $i = 1 To $aAppointments[0][0]
				If $aAppointments[$i][3] Then ; if appointment is happening
					startMessage($aAppointments[$i][0])
				EndIf
			Next
		EndIf
	EndIf
	
	$aLastAppointments = $aAppointments
	
	Sleep(250) ; reduce CPU usage
WEnd
