; Pseudo code for final SerialInput.au3
; Assumptions: SNs provided are correct.
;
; Read cleanSNs.txt and store SNs
; navigate to Receiving module
; enter order # and bill of lading
; select appropriate SKU and enter quantity
; enter SNs
; check all SNs were entered
; submit

#include <IE.au3>
#include <File.au3>

HotKeySet("^{SPACE}", "abort")		; CTRL + SPACEBAR to abort.

Local $path = "Path Removed."
Local $serialArray[200]

If Not _FileReadToArray($path,$serialArray,$FRTA_COUNT + $FRTA_INTARRAYS," ") Then
        MsgBox($MB_SYSTEMMODAL,"","There was an error reading the file." & _
        	" @error: " & @error)
EndIf

Sleep (3000)

; Uncomment the appropriate For loop.

; Receiving.
;#comments-start
For $vElement In $serialArray[1]
	MouseClick("left",83,265,2)
	Send($vElement)
	MouseClick("left",25,304,1)
	Sleep(2000)
Next
;#comments-end

#comments-start
; Cycle counts:
For $vElement In $serialArray[1]
	If StringLen($vElement) > 5 Then
		MouseClick("left",71,336,2)
		Send($vElement)
		MouseClick("left",262,433,1)
		Sleep(3000)
	EndIf
Next
#comments-end

Func abort()
	Exit
EndFunc
