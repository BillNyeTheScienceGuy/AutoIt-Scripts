#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         BillNyeTheScienceGuy

 Script Function:
	Displays Outlook calendar appointments on a Dream Cheeky LED message board 
	using the LED Panel Controller program made by Tiago Rodrigues.
	
	NOTE: Script must be in the same directory as the LED Panel Controller 
	executable file.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <Date.au3>      ; for _DateDiff and _DateAdd
#include <ExtMsgBox.au3> ; for _ExtMsgBoxSet and _ExtMsgBox
#include <Misc.au3>      ; for _IsPressed
#include <Outlook.au3>   ; for _OutlookOpen and _OutlookGetAppointments
#include <String.au3>    ; for _StringInsert

; Function Name..:  eventIsHappening()
; Description....:  Indicate if current time lies within event times
; Syntax.........:  eventIsHappening($sStart, $sEnd)
; Parameter(s)...:  $sStart - String of event start time, format YYYYMMDDHHMMSS
;                   $sEnd   - String of event end time, format YYYYMMDDHHMMSS
; Requirement(s).:  String and Date libraries
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

; Function Name..:  postMessage()
; Description....:  Scroll a message on the LED board and then delete it after it has passed once
; Syntax.........:  postMessage($sMessage)
; Parameter(s)...:  $sMessage - String of message to have scrolled across board
; Requirement(s).:  LEDDisplayControllerGui.exe must be in script directory and the server must currently be running
;                   LEDDisplayControllerGui.exe must be in the C: drive
; Return Value(s):  None
Func postMessage($sMessage)
	Local $iMillisecondsPerCharacter = 400
	Local $iCharactersInString = StringLen($sMessage)
	
	; -SetRepeatCount and -DeleteMsgAfterRepeat don't work, so the Sleep function and -DeleteRegex is necessary
	Run("LedDisplayControllerGui.exe -NoServer -NoGui -SetRepeatCount 1 -DeleteMsgAfterRepeat 1 -SendText """ & $sMessage & """")
	Local $iWaitTime = $iMillisecondsPerCharacter*$iCharactersInString
	Sleep($iWaitTime)
	Run("LedDisplayControllerGui.exe -NoServer -NoGui -DeleteRegex """"")
EndFunc

; Function Name..:  loadAppointments()
; Description....:  Open an outlook object and put the appointment details in the passed array
; Syntax.........:  loadAppointments($aAppointments)
; Parameter(s)...:  $aAppointments - Array to place appointment details in
; Requirement(s).:  Outlook UDF (https://www.autoitscript.com/forum/topic/89321-outlook-udf/) and Date libraries
; Return Value(s):  None
Func loadAppointments(ByRef $aAppointments)
	Local $oOutlook = _OutlookOpen()
	Local $sDate = @YEAR & "-" & @MON & "-" & @MDAY ; current date in YYYY-MM-DD format
	$aAppointments = _OutlookGetAppointments($oOutlook, "", $sDate & " 00:00", _DateAdd('d', 1, $sDate) & " 00:00", "", 2, "")
EndFunc

; Function Name..:  loadNewMessages()
; Description....:  Open an outlook object and put all unread mail details in the passed array
; Syntax.........:  loadNewMessages($aMessages)
; Parameter(s)...:  $aMessages - Array to place unread mail details in
; Requirement(s).:  Outlook UDF (https://www.autoitscript.com/forum/topic/89321-outlook-udf/)
; Return Value(s):  None
Func loadUnreadMessages(ByRef $aMessages)
	Local $oOutlook = _OutlookOpen()
	$aMessages = _OutlookGetMail($oOutlook, $olFolderInbox, False, "", "", "", "", "", "", True)
EndFunc

; Function Name..:  listNewMessages()
; Description....:  Display unread Outlook messages in a basic message box
; Syntax.........:  listNewMessages()
; Parameter(s)...:  None
; Requirement(s).:  None
; Return Value(s):  None
Func listUnreadMessages()
	Local $aMessages
	loadUnreadMessages($aMessages)
	
	If $aMessages[0][1] == 0 Then
		MsgBox(1, "Unread Mail", "No unread messages")
	Else
		Local $iLongestNameLength = 0
		Local $sMessage = "Unread mail:" & @LF
		
		; Find longest name length for formatting purposes
		For $i = 1 To $aMessages[0][1]
			If $iLongestNameLength < StringLen($aMessages[$i][0]) Then
				$iLongestNameLength = StringLen($aMessages[$i][0])
			EndIf
		Next
		
		; Adds together name and subject strings with space padding for different sized names
		For $i = 1 To $aMessages[0][1]
			$sMessage = $sMessage & @LF & "- " & $aMessages[$i][0] & _StringRepeat(" ", $iLongestNameLength - StringLen($aMessages[$i][0])) & " -> " & $aMessages[$i][7] ; add sender and subject to new line
		Next
		
		_ExtMsgBoxSet(-1, -1, -1, -1, -1, "Courier New")
		_ExtMsgBox(0, 0, "Unread Mail", $sMessage)
	EndIf
EndFunc

; Function Name..:  forceQuitAllProcesses()
; Description....:  Force quit all processes of a certain name
; Syntax.........:  forceQuitAllProcesses($sProcessName)
; Parameter(s)...:  $sProcessName - String of the process name to quit
; Requirement(s).:  None
; Return Value(s):  None
Func forceQuitAllProcesses($sProcessName)
	Local $aProcesses = ProcessList($sProcessName)
	For $i = 0 To UBound($aProcesses) - 1
		Run("taskkill /F /IM " & $sProcessName)
	Next
EndFunc

; Function Name..:  restartScript()
; Description....:  Runs the current version of this script and exits (https://gist.github.com/kissgyorgy/4350758)
; Syntax.........:  restartScript()
; Parameter(s)...:  None
; Requirement(s).:  This script should be in the C: drive so the command line can reach it
; Return Value(s):  None
Func restartScript()
    If @Compiled = 1 Then
        Run(FileGetShortName(@ScriptFullPath))
    Else
        Run(FileGetShortName(@AutoItExe) & " " & FileGetShortName(@ScriptFullPath))
    EndIf
	
    Exit
EndFunc

Local $aAppointments;, $aMessages, $iNumberOfUnreadMessages
Local $iMessageGapTime = 1000 ; in milliseconds

HotKeySet("^!u", "listUnreadMessages") ; set Ctrl + Alt + U as the hotkey to list unread messages
HotKeySet("^!r", "restartScript") ; set Ctrl + Alt + R as the hotkey to restart the script

; Quit all previous controller processes, then start LED Controller app server with no GUI
forceQuitAllProcesses("LedDisplayControllerGui.exe")
Sleep(500) ; give process-quiting time to stop so a new process can be created
Run("LedDisplayControllerGui.exe -NoGui") ; start controller server with no GUI

; Loops through the appointments and displays all the ones currently taking place
While True
	
	; loadUnreadMessages($aMessages)
	
	; If $iNumberOfUnreadMessages < $aMessages[0][1] Then
		; $iNumberOfUnreadMessages = $aMessages[0][1]
		; listUnreadMessages()
	; Else
		; $iNumberOfUnreadMessages = $aMessages[0][1]
	; EndIf
	
	loadAppointments($aAppointments)
	
	If IsArray($aAppointments) Then ; if appointments exist
		For $i = 1 To $aAppointments[0][0]
			If eventIsHappening($aAppointments[$i][1], $aAppointments[$i][2]) Then
				postMessage($aAppointments[$i][0])
				Sleep($iMessageGapTime)
			EndIf
		Next
	EndIf
	
	Sleep(250) ; reduce CPU usage
WEnd

; TODO: list unread messages whenever the number of unread messages increases
;       prevent this constant checking from interrupting taskbar thumbnail preview
