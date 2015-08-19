#include <IE.au3>
#include <Excel.au3>
#include "CloseHiddenIE.au3"

Global $oBrowser

closeHiddenIE()
getReport()


Func getReport()
	Local $oUsername, $oPassword, $oLogin, $oOrderStatus, $oFacility, _
		  $oSubmit, $oExcel, $hPopUp
	Local $sUrl = "URL Removed."

	TrayTip("Open Orders", "Navigating to report...", 600)
	$oBrowser = _IECreate("URL Removed.", 0, 0, 1, 0)

	If @error <> 0 Then
		MsgBox("", "Open Orders Script Fatal Error", "Error code: " & _
			@error & "." & @CRLF & @CRLF & "Please restart the script.")
		_IEQuit($oBrowser)
		Exit
	EndIf

	$oUsername = _IEGetObjById($oBrowser, "username")
	$oPassword = _IEGetObjById($oBrowser, "password")
	$oLogin = _IEGetObjById($oBrowser, "IMAGE1")
	_IEFormElementSetValue($oUsername, "Removed.")
	_IEFormElementSetValue($oPassword, "Removed.")
	_IEAction($oLogin, "click")
	_IELoadWait($oBrowser)

	_IENavigate($oBrowser, $sUrl, 0)
	TrayTip("Open Orders", "Loading report...", 600)
	timeOutNotice($oBrowser)
	$oOrderStatus = _IEGetObjById($oBrowser, "RP1_1")
	$oFacility = _IEGetObjById($oBrowser, "RP1_2")
	$oSubmit = _IEGetObjById($oBrowser, "Submit")

	; Select all options except 'CANCELED' & 'CANCELLED'.
	For $i = 2 to 20 Step +1
		; Except 'SHIPED' & 'SHIPPED'.
		If ($i <> 18) And ($i <> 19) Then
			_IEFormElementOptionSelect($oOrderStatus, $i, 1, "byIndex")
		EndIf
	Next

	; Select all options except all T5500 & T8000 plants.
	For $i = 0 to 30 Step +1
		; Except T1015.
		If ($i <> 8) Then
			_IEFormElementOptionSelect($oFacility, $i, 1, "byIndex")
		EndIf
	Next

	For $i = 52 to 78 Step +1
		_IEFormElementOptionSelect($oFacility, $i, 1, "byIndex")
	Next

	_IEAction($oSubmit, "click")
	timeOutNotice($oBrowser)

	$oExcel = _IEGetObjById($oBrowser, "imgNativeExcel")
	_IEAction($oExcel, "click")
	TrayTip("Open Orders", "Downloading report...", 600)
	$hPopUp = WinWait("Please Wait - Windows Internet Explorer")
	WinSetState($hPopUp, "", @SW_HIDE)
	_IEQuit($oBrowser)
	WinWait("[REGEXPTITLE:Microsoft Excel - \S{24}-\d{9,10}.xls]", "")
	WinClose($hPopUp)
EndFunc

Func timeOutNotice($oBrowser)

	_IELoadWait($oBrowser, 0, 600000)
	If @error == 6 Then
		MsgBox("", "Open Orders Report Timed Out", "LOGI hasn't responded " & _
				"in 10 minutes. Exiting script.")
		_IEQuit($oBrowser)
		Exit
	EndIf
EndFunc
