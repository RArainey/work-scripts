#include <IE.au3>
#include <Excel.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include "YearsFromToday.au3"
#include "CloseHiddenIE.au3"

Local $aPlantCode[ 31 ] = [ "Actual plant codes were removed." ], _
	  $idDropDown, $idRun, _
	  $sList = ""

; Copy plant codes from array to string as required by drop down GUI.

For $i = 0 To ( UBound( $aPlantCode ) - 1 ) Step +1
    $sList &= "|" & $aPlantCode[ $i ]
Next

closeHiddenIE()

#Region Create GUI

GUICreate( "Inbound Orders", 200, 50 )

GUICtrlCreateLabel( "Select plant:", 10, 10 )

$idDropDown = GUICtrlCreateCombo( "", 10, 25 )
GUICtrlSetData( $idDropDown, $sList )

$idRun = GUICtrlCreateButton( "Print", 110, 15, 80 )

GUISetState()

#EndRegion Create GUI

#Region Execution

While 1 ; Run until user exits.
  Switch GUIGetMsg()
      Case $GUI_EVENT_CLOSE
          Exit

	Case $idRun
		InboundOrders()
		FormatReport()
		Exit
  EndSwitch
WEnd

#EndRegion Execution

#Region Functions
; *****************************************************************************
; InboundOrders()
; Logs into website and downloads PO/STO report in .csv format.
; Returns nothing.
; *****************************************************************************
Func InboundOrders()

	GUISetState( @SW_HIDE )

	Local $oBrowser, $hNewIE, $oUsername, $oPassword, $oLogin, $oOrderStatus, _
		  $oPlant, $oStartDate, $oEndDate, $oSubmit, $oCsv, _
		  $sPlantCode = GUICtrlRead( $idDropDown ), _
		  $sPastDate =  yearsFromToday( -1 ), _
		  $sFutureDate = yearsFromToday( 1 )

	TrayTip( "Inbound Orders", "Navigating to PO/STO report...", 5 )

	$oBrowser = _IECreate( "website URL", 0, 0 )

	$oUsername = _IEGetObjById( $oBrowser, "username" )
	$oPassword = _IEGetObjById( $oBrowser, "password" )
	$oLogin = _IEGetObjById( $oBrowser, "IMAGE1" )
	_IEFormElementSetValue( $oUsername, "removed" )
	_IEFormElementSetValue( $oPassword, "removed" )
	_IEAction( $oLogin, "click" )

	_IELoadWait( $oBrowser )
	_IENavigate( $oBrowser, "website URL" )

	$oOrderStatus = _IEGetObjById( $oBrowser, "RP1_1" )
	$oPlant = _IEGetObjById( $oBrowser, "RP1_2" )
	$oStartDate = _IEGetObjById( $oBrowser, "RP1_5" )
	$oEndDate = _IEGetObjById( $oBrowser, "RP1_6" )
	$oSubmit = _IEGetObjById( $oBrowser, "Submit" )

	; Change report parameters and 'run'.
	_IEFormElementSetValue( $oOrderStatus, "OPEN" )
	_IEFormElementSetValue( $oPlant, $sPlantCode )
	_IEFormElementSetValue( $oStartDate, $sPastDate )
	_IEFormElementSetValue( $oEndDate, $sFutureDate )
	_IEAction( $oSubmit, "click" )

	TrayTip( "Inbound Orders", "Downloading PO/STO report...", 5 )

	; Assign newly available HTML element to variable and download CSV.
	_IELoadWait( $oBrowser )
	$oCsv = _IEGetObjById( $oBrowser, "CSV" )
	_IEAction( $oCsv, "click" )

	$hNewIE = WinWait( "Please Wait - Windows Internet Explorer" )
	WinSetState( $hNewIE, "", @SW_HIDE )
	_IEQuit( $oBrowser )

	WinWait( "[REGEXPTITLE:Microsoft Excel - \S{24}-\d{9,10}.csv]", "" )
	Winclose( $hNewIE )

EndFunc

; *****************************************************************************
; FormatReport()
; Reorders columns and adjusts visual style of PO/STO report before printing.
; *****************************************************************************
Func FormatReport()

	Local $oExcel, $oRange, $oReport = ""

	Local $sPattern = '\S{24}-\d{9,10}\.csv'

	$oExcel = ObjGet( "", "Excel.Application" )
	If @error Then MsgBox( "", "Excel ObjGet Failed", "@error = " & @error )

	; Find the PO/STO workbook and attach it to a variable.
	For $oWB in $oExcel.Workbooks
		If StringRegExp( $oWB.Name, $sPattern ) Then
			$oReport = $oWB
		EndIf
	Next

	If StringCompare( $oReport, "" ) Then
		MsgBox( "" , "FormatReport() Failure", "Workbook not found," & _
		  aborting script." )
		Exit
	EndIf

	_Excel_RangeCopyPaste( $oReport.ActiveSheet, "B:B", "A:A" )
	_Excel_RangeCopyPaste( $oReport.ActiveSheet, "C:C", "B:B" )
	_Excel_RangeCopyPaste( $oReport.ActiveSheet, "E:E", "C:C" )
	_Excel_RangeCopyPaste( $oReport.ActiveSheet, "F:F", "D:D" )
	_Excel_RangeCopyPaste( $oReport.ActiveSheet, "I:I", "E:E" )
	_Excel_RangeCopyPaste( $oReport.ActiveSheet, "K:K", "F:F" )

	_Excel_RangeDelete( $oReport.ActiveSheet, "G:R" )

	_Excel_RangeSort( $oReport, $oReport.ActiveSheet, $oReport.ActiveSheet.UsedRange, _
					"A:A", Default, Default, $xlYes )

	$oReport.ActiveSheet.Columns( "A:F" ).EntireColumn.AutoFit

	; xlLineStyle Constants
	$xlContinuous = 1
	$xlDash = -4115
	$xlDashDot = 4
	$xlDashDotDot = 5
	$xlDot = -4118
	$xlDouble = -4119
	$xlLineStyleNone = -4142
	$xlSlantDashDot = 13

	; XlBordersIndex Constants
	$xlDiagonalDown = 5
	$xlDiagonalUp = 6
	$xlEdgeBottom = 9
	$xlEdgeLeft = 7
	$xlEdgeRight = 10
	$xlEdgeTop = 8
	$xlInsideHorizontal = 12
	$xlInsideVertical = 11

	With $oReport.ActiveSheet.UsedRange
		.Borders( $xlEdgeBottom ).LineStyle = $xlContinuous
		.Borders( $xlEdgeTop ).LineStyle = $xlContinuous
		.Borders( $xlEdgeLeft ).LineStyle = $xlContinuous
		.Borders( $xlEdgeRight ).LineStyle = $xlContinuous
		.Borders( $xlInsideHorizontal ).LineStyle = $xlContinuous
		.Borders( $xlInsideVertical ).LineStyle = $xlContinuous
	EndWith

	$oReport.ActiveSheet.Range( "A1:F1" ).Font.Bold = True

	_Excel_Print( $oExcel, $oReport.ActiveSheet.UsedRange )

	If @error Then
		Exit
		MsgBox( $MB_SYSTEMMODAL, "PO/STO Report Failed To Print.", "Is the " & _
				"correct printer set to default?" & @CRLF & "@error code: " & _
				@error & ". @extend: " & @extended & "." )
	EndIf
EndFunc
#EndRegion Functions
