#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Robert Spiller

 Script Function:
	Displays Outlook calendar appointments on a Dream Cheeky LED message board 
	using the LED Panel Controller program made by Tiago Rodrigues.
	
	NOTE: Script must be in the same directory as the LED Panel Controller 
	executable file.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <Date.au3>    ; for _DateDiff and _DateAdd
#include <Outlook.au3> ; for _OutlookOpen and _OutlookGetAppointments
#include <String.au3>  ; for _StringInsert

; Function Name:    eventIsHappening()
; Description:      Indicate if current time lies within event times
; Syntax.........:  eventIsHappening($sStart, $sEnd)
; Parameter(s):     $sStart - String of event start time, format YYYYMMDDHHMMSS
;                   $sEnd   - String of event end time, format YYYYMMDDHHMMSS
; Requirement(s):   String and Date libraries
; Return Value(s):  On Success - Returns True
;                   On Failure - Returns False
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
	
	; MsgBox(1, "", $sStart & ", " & $sEnd & "; " & $bAtOrAfterStart & ", " & $bAtOrBeforeEnd)
	
	; Return True iff current time is [start, end]
	If $bAtOrAfterStart And $bAtOrBeforeEnd Then
		Return True
	EndIf
	Return False
EndFunc

; Function Name:    postMessage()
; Description:      Scroll a message on the LED board and then delete it after it has passed once
; Syntax.........:  postMessage($sMessage)
; Parameter(s):     $sMessage - String of message to have scrolled across board
; Requirement(s):   LEDDisplayControllerGui.exe must be in script directory and the server must currently be running
;                   LEDDisplayControllerGui.exe must be in the C: drive
; Return Value(s):  None
Func postMessage($sMessage)
	Local $iMillisecondsPerCharacter = 420
	Local $iCharactersInString = StringLen($sMessage)
	
	; -SetRepeatCount and -DeleteMsgAfterRepeat don't work, so the Sleep function and -DeleteRegex is necessary
	Run("LedDisplayControllerGui.exe -NoServer -NoGui -SetRepeatCount 1 -DeleteMsgAfterRepeat 1 -SendText """ & $sMessage & """")
	Local $iWaitTime = $iMillisecondsPerCharacter*$iCharactersInString
	Sleep($iWaitTime)
	Run("LedDisplayControllerGui.exe -NoServer -NoGui -DeleteRegex """"")
EndFunc

; Function Name:    loadAppointments()
; Description:      Open an outlook object and put the appointment details in the passed array
; Syntax.........:  loadAppointments($aAppointments)
; Parameter(s):     $aAppointments - Array to place appointment details in
; Requirement(s):   Outlook UDF (https://www.autoitscript.com/forum/topic/89321-outlook-udf/) and Date libraries
; Return Value(s):  None
Func loadAppointments(ByRef $aAppointments)
	Local $oOutlook = _OutlookOpen()
	Local $sDate = @YEAR & "-" & @MON & "-" & @MDAY ; current date in YYYY-MM-DD format
	$aAppointments = _OutlookGetAppointments($oOutlook, "", $sDate & " 00:00", _DateAdd('d', 1, $sDate) & " 00:00", "", 2, "")
EndFunc

Local $aAppointments
Local $iGapTime = 1000 ; in milliseconds

; Loops through the appointments and displays all the ones currently taking place
While True
	loadAppointments($aAppointments)
	For $i = 1 To $aAppointments[0][0]
		If eventIsHappening($aAppointments[$i][1], $aAppointments[$i][2]) Then
			postMessage($aAppointments[$i][0])
			Sleep($iGapTime)
		EndIf
	Next
	Sleep(1000) ; reduce CPU usage
WEnd

; TODO: load appointments less frequently (maybe every 10 seconds) to allow windows taskbar to function without clicking it
;       could possibly require two scripts that share variables within a common GUI
