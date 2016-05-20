#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         BillNyeTheScienceGuy

 Script Function:
	Displays Outlook unread messages in a pop-up message box.  The box is forced 
	to display on top of all other windows.  The message box displays whenever 
	the number of unread messages increases.  The message box will also dispaly 
	when the key bind ctrl + alt + u is used.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <ExtMsgBox.au3> ; for _ExtMsgBoxSet and _ExtMsgBox
#include <Notify.au3>    ; for _Notify_Locate, _Notify_Set, and _Notify_Show
#include <Outlook.au3>   ; for _OutlookOpen and OutlookGetMail
#include <String.au3>    ; for _StringRepeat

; Function Name..:  loadUnreadMessages()
; Description....:  Open an outlook object and put all unread mail details in the passed array
; Syntax.........:  loadUnreadMessages($aMessages)
; Parameter(s)...:  $aMessages - Array to place unread mail details in
; Requirement(s).:  Outlook UDF (https://www.autoitscript.com/forum/topic/89321-outlook-udf/)
; Return Value(s):  None
Func loadUnreadMessages(ByRef $aMessages)
	Local $oOutlook = _OutlookOpen()
	$aMessages = _OutlookGetMail($oOutlook, $olFolderInbox, False, "", "", "", "", "", "", True)
EndFunc

; Function Name..:  listUnreadMessages()
; Description....:  Display unread Outlook messages in a basic message box
; Syntax.........:  listUnreadMessages()
; Parameter(s)...:  None
; Requirement(s).:  ExtMsgBox UDF (https://www.autoitscript.com/forum/topic/109096-extended-message-box-bugfix-version-9-aug-15/)
;                   String library
; Return Value(s):  None
Func listUnreadMessages()
	Local $aMessages
	Local $sMessage = "Unread mail:" & @LF
	loadUnreadMessages($aMessages)
	
	If $aMessages[0][1] == 0 Then
		$sMessage &= @LF & "No unread messages"
	Else
		Local $iLongestNameLength = 0
		
		; Find longest name length for formatting purposes
		For $i = 1 To $aMessages[0][1]
			If $iLongestNameLength < StringLen($aMessages[$i][0]) Then
				$iLongestNameLength = StringLen($aMessages[$i][0])
			EndIf
		Next
		
		; Adds together name and subject strings with space padding for different sized names
		For $i = 1 To $aMessages[0][1]
			$sMessage &= @LF & "- " & $aMessages[$i][0] & _StringRepeat(" ", $iLongestNameLength - StringLen($aMessages[$i][0])) & " -> " & $aMessages[$i][7]
		Next
	EndIf
	
	_ExtMsgBoxSet(-1, -1, -1, -1, -1, "Courier New")
	_ExtMsgBox(0, 0, "Unread Mail", $sMessage)
	WinSetOnTop("Unread Mail", $sMessage, 1)
	; _Notify_Locate()
	; _Notify_Set(0, -1, -1, "Courier New")
	; _Notify_Show(0, "Unread Mail", $sMessage)
EndFunc

HotKeySet("^!u", "listUnreadMessages") ; set Ctrl + Alt + U as the hotkey to list unread messages

Local $aMessages, $iNumberOfUnreadMessages

; Loops through the appointments and displays all the ones currently taking place
While True
	loadUnreadMessages($aMessages)
	
	If $iNumberOfUnreadMessages < $aMessages[0][1] Then
		$iNumberOfUnreadMessages = $aMessages[0][1]
		listUnreadMessages()
	Else
		$iNumberOfUnreadMessages = $aMessages[0][1]
	EndIf
	
	Sleep(250) ; reduce CPU usage
WEnd
