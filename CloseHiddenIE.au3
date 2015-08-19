#include <IE.au3>

; *****************************************************************************
; closeHiddenIE()
; Closes all instances of Internet Explorer that are invisible.
; *****************************************************************************

Func closeHiddenIE()
	; Retrieve a list of IE window handles.
	Local $aList = WinList("[CLASS:IEFrame]")

	; Loop through and close all hidden IE windows.
	For $i = 1 To $aList[0][0]
		If ( $aList[$i][0] <> "" ) And _
				Not ( BitAND(WinGetState($aList[$i][1]), 2) ) Then
			WinClose( $aList[$i][1] )
		EndIf
	Next
EndFunc
