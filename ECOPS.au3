
#include <IE.au3>
#include <Date.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ColorConstants.au3>
#include "YearsFromToday.au3"
#include "CloseHiddenIE.au3"

#Region Execution

Global $oBrowser, $hBrowser, $sUsername, $sPassword, $sFailedReceive, _
	   $sNoRecord, $aPO, $bFailedPO = False, _
	   $aUrl[5] = [ "URLS Removed." ]
Local $sPOStatus = "Successfully submitted these POs:" & @CRLF

closeHiddenIE()
captureData()
intoEtos()

For $sPO in $aPO
	findPackingSlip($sPO, False)
Next

_IEQuit($oBrowser)

If $bFailedPO Then
	If ($sNoRecord <> "") Then
		$sPOStatus &= @CRLF & @CRLF & "No receipt report was found for:" & _
					@CRLF & $sNoRecord & @CRLF & @CRLF & "Please" & _
					" check the numbers are correct."
	EndIf
	
	If ($sFailedReceive <> "") Then
		$sPOStatus &= @CRLF & @CRLF & "These POs failed to receive:" & _
					@CRLF & $sFailedReceive
	EndIf
EndIf

MsgBox("", "Auto-ECOPS", $sPOStatus)

#EndRegion Execution

#Region Functions


; *****************************************************************************
; captureData()
; Creates a GUI where users enter their ID, password, and PO numbers.
; captureData() runs until either the user quits or their provided data passes
; validation.
; *****************************************************************************
Func captureData()

	Global $idUsername, $idPassword, $idPO
	Local $idReceieve

	GUICreate("Auto-ECOPS", 190, 290)
	GUICtrlCreateLabel("ID:", 10, 10)
	GUICtrlSetFont(-1, 9, 700)
	$idUsername = GUICtrlCreateInput("", 80, 8, 90)
	GUICtrlSetLimit(-1, 7)
	GUICtrlCreateLabel("Password::", 10, 40)	; 1 colon doesn't show up...
	GUICtrlSetFont(-1, 9, 700)
	$idPassword = GUICtrlCreateInput("", 80, 38, 90, Default, $ES_PASSWORD)
	GUICtrlCreateLabel("ECOPS POs:", 10, 73)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("    e.g. 4501112222,4502223333" & @CRLF & "   " & _
		"                No spaces!", 10, 93)
	$idPO = GUICtrlCreateEdit("", 20, 123, 150, 125, $ES_MULTILINE)
	GUICtrlSetTip(-1, "e.g. 4501112222,4502223333,4503334444")
	$idReceieve = GUICtrlCreateButton("Receive Orders", 55, 255)

	GUISetState(@SW_SHOW)

	; Loop until the user exits or successfully enters their PO numbers.
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				Exit	; User closed the window.
			Case $idReceieve
				If validate() Then
					GUIDelete()
					Return
				EndIf
		EndSwitch
	WEnd

EndFunc   ; End captureData()


; *****************************************************************************
; validate()
; Checks username, password, and PO numbers are correct.
; Returns boolean.
; *****************************************************************************
Func validate()

	$sUsername = GUICtrlRead($idUsername)
	$sPassword = GUICtrlRead($idPassword)
	Local $sRawPO = GUICtrlRead($idPO)
	Local $bUsername = False, $bPassword = False, $bPO = False
	Local $sError = "The PO numbers are invalid." & @CRLF &  "Check that the" & _
					"PO numbers start with 450 and are 10-digits long." & @CRLF & _
					@CRLF & "e.g. 4501112222,4502223333"
	Local $iLength

	If Not StringRegExp($sUsername, '(?i)t\d{6}') Then
		GUICtrlSetBkColor($idUsername, $COLOR_RED)
		GUICtrlSetFont($idUsername, Default, 700)
		$bUsername = False
	Else
		GUICtrlSetBkColor($idUsername, Default)
		GUICtrlSetFont($idUsername, Default, Default)
		$bUsername = True
	EndIf

	If $sPassword == "" Then
		GUICtrlSetBkColor($idPassword, $COLOR_RED)
		$bPassword = False
	Else
		GUICtrlSetBkColor($idPassword, Default)
		$bPassword = True
	EndIf

	If StringLen($sRawPO) > 10 Then

		; Search for every char except digits and commas.
		If (StringRegExp($sRawPO, '[^\d,]')) Then
			MsgBox("", "PO error", $sError)

			GUICtrlSetBkColor($idPO, $COLOR_RED)
			GUICtrlSetFont($idPO, Default, 700)
			$bPO = False

			; Quit before checking the POs further.
			Return False
		Else

			GUICtrlSetBkColor($idPO, Default)
			GUICtrlSetFont($idPO, Default, Default)
			$bPO = True
		EndIf
	EndIf

	$aPO = StringSplit($sRawPO, ",", "2")
	$iLength = UBound($aPO, 1) - 1	; Array length. 0 indexed.

	For $i = 0 to $iLength Step +1
		If (Not (StringRegExp($aPO[$i], '450\d{7}')) Or _
			  (StringLen($aPO[$i]) <> 10)) Then

			MsgBox("", "PO error", $sError)
			GUICtrlSetBkColor($idPO, $COLOR_RED)
			GUICtrlSetFont($idPO, Default, 700)
			$bPO = False
		Else
			GUICtrlSetBkColor($idPO, Default)
			GUICtrlSetFont($idPO, Default, Default)
			$bPO = True
		EndIf
	Next

	If ($bUsername And $bPassword And $bPO) Then
		Return True
	EndIf

EndFunc	; End validate()


; *****************************************************************************
; intoEtos()
; Opens a new instance of IE and logs in to eTOS. Prompts user for credentials
; if bad eTOS username/password.
; *****************************************************************************
Func intoEtos()

	Local $oUsername, $oPassword, $oSubmit
	$oBrowser = _IECreate($aUrl[0], 0, 0)
	$hBrowser = _IEPropertyGet($oBrowser, "hwnd")
	$oUsername = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_username")
	$oPassword = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_pwd")
	$oSubmit = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_Button1")
	_IEFormElementSetValue($oUsername, "Removed.")
	_IEFormElementSetValue($oPassword, "Removed.")
	_IEAction($oSubmit, "click")
	_IELoadWait($oBrowser)
	
	$oSubmit = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_imgEnglish")
	_IEAction($oSubmit, "click")
	_IELoadWait($oBrowser)
	
	$oUsername = _IEGetObjById($oBrowser, "T1")
	$oPassword = _IEGetObjById($oBrowser, "T2")
	$oSubmit = _IEGetObjByName($oBrowser, "B1")
	_IEFormElementSetValue($oUsername, $sUsername)
	_IEFormElementSetValue($oPassword, $sPassword)
	_IEAction($oSubmit, "click")
	_IELoadWait($oBrowser)

	; Make IE visible and have the user log in if their eTOS credentials failed.
	If StringRegExp(String(_IEPropertyGet($oBrowser, "locationurl")), _
					'' & String($aUrl[1]) & '') Then
		WinSetState($hBrowser, "", @SW_SHOW)
		WinActivate($hBrowser)
		MsgBox("", "Auto-ECOPS", "The eTOS username and password failed." & _
				@CRLF & "Please log into eTOS to continue.")

		; Keep flashing until the user logs into eTOS.
		While StringRegExp(String(_IEPropertyGet($oBrowser, "locationurl")), _
			'' & String($aUrl[1]) & '')
			WinFlash($hBrowser)
		WEnd
	EndIf

	; Hide the IE window if it's visible.
	If BitAND(WinGetState($hBrowser), 2) Then
		WinSetState($hBrowser, "", @SW_HIDE)
	EndIf
EndFunc


; *****************************************************************************
; findPackingSlip()
; Retrieves packing slip and saves SKU and QTY to 2d array[row][column]
; then calls dataEntry. If it fails to retrieve a packing slip it saves the PO #
; and alerts the user at the end of execution.
; findPackingSlip checks if an order was received if $bConfirmReceived is
; passed in True.
; *****************************************************************************
Func findPackingSlip($sPO, $bConfirmReceived)

	Local $aPackingSlip, $oPlant, $oPo, $oRun, $oTable, $oDateFrom, $oDateTo

	If ($bConfirmReceived == False) Then
		TrayTip("Auto-ECOPS", "Now processing PO " & $sPO & ".", 5, 1)
	EndIf

	_IENavigate($oBrowser, $aUrl[3])
	$oPlant = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_ddlPlant")
	$oDateFrom = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbDateFrom")
	$oDateTo = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbDateTo")
	$oPo = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbDocDaw")
	$oRun = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_btnContinue")
	_IEFormElementOptionSelect($oPlant, "1101")
	_IEFormElementSetValue($oDateFrom, yearsFromToday(-1))
	_IEFormElementSetValue($oDateTo, yearsFromToday(1))
	_IEFormElementSetValue($oPo, $sPO)
	_IEAction($oRun, "click")
	_IELoadWait($oBrowser)

	If $bConfirmReceived Then
		$sText = _IEBodyReadText($oBrowser)

		If StringInStr($sText, "CLOSED") Then
			$sPOStatus &= @CRLF & "	" & $sPO
			$bConfirmReceived = False
			Return
		ElseIf StringInStr($sText, "OPEN") Then
			$sFailedReceive &= $sPO & @CRLF
			$bFailedPO = True
			Return
		Else
			$sFailedReceive &= $sPO & @CRLF
			$bFailedPO = True
			Return
		EndIf
	EndIf

	$oTable = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_gvDetailDTL")

	If (@error <> 0) And (Not $bConfirmReceived) Then
		$sNoRecord &= $sPO & @CRLF
		$bFailedPO = True
	Else
		$aPackingSlip = _IETableWriteToArray($oTable, True)

		; Note: When working with $aPackingSlip we really only need column 4:
		; Part Number and column 7: Quantity.
		; e.g. Line 1, part number -> $aPackingSlip[1][4]
		; 	   Line 3, quantity -> $aPackingSlip[3][7]

		dataEntry($aPackingSlip)
	EndIf
EndFunc


; *****************************************************************************
; dataEntry()
; Navigates to Minor Material Tracker and enters ECOPS packing slip details.
; It makes a findPackingSlip call to verify if the PO was received.
;
; Note: This webpage validates input onChange and then refreshes. All HTML
; elements assigned to objects will be lost during a refresh. Assign objects
; immediately before interaction or they will not work!
; *****************************************************************************
Func dataEntry($aPackingSlip)

	Local $oContinue, $oDropdown, $oPlant, $oNotes, $oMaterial, $oQuantity, _
		  $oStorage, $oValuation, $oReel, $oPoNbr, $oBOL, $oPo, $oLineNumber, _
		  $oFrame, $sText, $sPartNumber, $iQuantity, _
		  $iTime = 0, _
		  $iLength = UBound($aPackingSlip, $UBOUND_ROWS) - 1, _
		  $sPO = String($aPackingSlip[1][1])

	_IENavigate($oBrowser, $aUrl[4])

	$oDropdown = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_ddlAction")
	$oContinue = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_imgContinue")

	_IEFormElementOptionSelect($oDropdown,"ADDMAT")
	_IEAction($oContinue, "click")
	_IELoadWait($oBrowser)

	$oDropdown = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_ddlDefaultlocation")
	$oNotes = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbComment")
	$oContinue = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_imgContinue")

	_IEFormElementOptionSelect($oDropdown, "KINSWUD1101")
	_IEFormElementSetValue($oNotes, "PO: " & $sPO)
	_IEAction($oContinue, "click")
	_IELoadWait($oBrowser)

	For $iRow = 1 to $iLength Step +1	; Skip the headers row.

		$sPartNumber = StringTrimRight(String($aPackingSlip[$iRow][4]), 1)
		$iQuantity = Int($aPackingSlip[$iRow][7])

		$oMaterial = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbPartNumber")
		_IEFormElementSetValue($oMaterial, $sPartNumber)

		Sleep ($iTime)
		_IELoadWait($oBrowser)

		$oQuantity = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbPartQty")
		_IEFormElementSetValue($oQuantity, $iQuantity)

		$oStorage =  _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_ddlStorageLoc")
		_IEFormElementOptionSelect($oStorage, "NEW", 1, "byText")

		; Check for part description. This determines how to handle Valuation & Reel.
		$oPartDescription = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_lbPartDESC")
		$sPartDescription = _IEPropertyGet($oPartDescription, "innertext")

		If (StringCompare("No Part Description", $sPartDescription) <> 0) Then

			; Part recognized. Enter Valuation, Reel will populate automatically.
			$oValuation = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_ddlValuationType")
			_IEFormElementOptionSelect($oValuation, "NEW")
			Sleep ($iTime)
			_IELoadWait($oBrowser)
		Else
			; Skip Valuation and enter 'NA' for Reel.
			$oReel = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbValutionType")
			_IEFormElementSetValue($oReel, "NA")
		EndIf

		$oPoNbr =  _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_rbPONbr")
		_IEAction($oPoNbr, "click")
		Sleep ($iTime)
		_IELoadWait($oBrowser)

		$oBOL = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbBillOfLading")
		_IEFormElementSetValue($oBOL, "NA")

		$oPo = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbPONbr")
		_IEFormElementSetValue($oPo, $sPO)

		$oLineNumber = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_tbLineNbr")
		_IEFormElementSetValue($oLineNumber, $iRow)

		$oContinue =  _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_imgContinue")
		_IEAction($oContinue, "click")
		Sleep ($iTime)
		_IELoadWait($oBrowser)

		; Check for error messages.
		$sText = _IEBodyReadText($oBrowser)

		If StringInStr($sText, "Blank/Space not Allowed in the quantity field") Or _
			StringInStr($sText, "Please select storage location from the drop down.") Or _
			StringInStr($sText, "Please enter your PO/STO/Delivery Number.") Or _
			StringInStr($sText, "Please enter Line Number.") Then

				; Re-enter this row, but slow everything down.
				$iRow = $iRow - 1
				$iTime = $iTime + 500
		EndIf
	Next	; end For

	TrayTip("Auto-ECOPS","Submitting PO " & $sPO & ".", 5)
	$oSubmit = _IEGetObjById($oBrowser, "ctl00_ContentPlaceHolder1_imgSubmit")
	_IEAction($oSubmit, "click")
	_IELoadWait($oBrowser)

	findPackingSlip($sPO, True)
EndFunc

#EndRegion Functions
