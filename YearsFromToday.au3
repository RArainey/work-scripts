#include <Date.au3>

; *****************************************************************************
; YearsFromToday()
; Returns the current date with adjusted year, +/- $iDifference, in
; MM/DD/YYYY format.
; *****************************************************************************
Func yearsFromToday( $iDifference )

	Local $aDateSplit, $aTimeSplit, $sResult = ""

	_DateTimeSplit( _NowCalcDate(), $aDateSplit, $aTimeSplit )

	$sResult = $aDateSplit[2] & "/" & $aDateSplit[0] & "/" & ( Int( $aDateSplit[1] ) + $iDifference )

	Return $sResult
EndFunc
