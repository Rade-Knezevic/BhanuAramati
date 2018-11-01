#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Include <WinAPI.au3>

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Form1", 238, 251, -1, -1)
$mPosX = GUICtrlCreateLabel(MouseGetPos()[0], 88, 112, 86, 17)
$LabelX = GUICtrlCreateLabel("Mouse X", 40, 112, 46, 17)
$mPosY = GUICtrlCreateLabel(MouseGetPos()[1], 88, 136, 86, 17)
$LabelY = GUICtrlCreateLabel("Mouse Y", 40, 136, 46, 17)
$mCursorDisplay = GUICtrlCreateLabel(getCursorInfo(), 88, 160, 86, 17)
$LabelY = GUICtrlCreateLabel("Cursor value", 40, 160, 46, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
	EndSwitch
	updateUI()
WEnd

Func updateUI()
	GUICtrlSetData($mPosX, MouseGetPos()[0])
    GUICtrlSetData($mPosY, MouseGetPos()[1])
	GUICtrlSetData($mCursorDisplay,getCursorInfo())
EndFunc

Func getCursorInfo()
	Global $cursor = _WinAPI_GetCursorInfo()
	$cursorLook =$cursor[2]
	return $cursorLook
EndFunc