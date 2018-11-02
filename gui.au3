#include <Array.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>
#include <MouseOnEvent.au3>

Global $Files[0]
Global $File_Attr[0][3]

Global $pathToDB = @ScriptDir & "\" &"Database4.accdb"
Global $pathToFiles = @ScriptDir & "\" &"files"
Global $pathToDownloads = @UserProfileDir & "\Downloads"

Global $Table_Name = "bhanu"
Global $Attr_Name[3] = ["", "", ""]

Global $FileToBeDownloaded
_DBUpdate()

$GroupHeight = 1
$GroupWidth = 1
$GroupsFromTop = 160
$FilesGroupFromLeft = 23
$RatingsGroupFromLeft = 123

$ListWidth = 95
$ListHeight = 195
$ListsFromTop = 180
$FileListFromLeft = 27
$RatingsListFromLeft = 125

$Form_Main = GUICreate("GUI managing Database", 250, 380)
$Group_Attributes = GUICtrlCreateGroup("Attributes", 20, 20, 200, 130)
$Checkbox_1 = GUICtrlCreateCheckbox($Attr_Name[0], 40, 50)
$Checkbox_2 = GUICtrlCreateCheckbox($Attr_Name[1], 40, 80)
$Checkbox_3 = GUICtrlCreateCheckbox($Attr_Name[2], 40, 110)
$idAddFile = GUICtrlCreateButton("Add", 140, 70, 60, 20)
$idDownloadFile = GUICtrlCreateButton("Download", 140, 110, 60, 20)
$Group_Files = GUICtrlCreateGroup("Files", $FilesGroupFromLeft, $GroupsFromTop, $GroupWidth ,$GroupHeight )
$Group_Ratings = GUICtrlCreateGroup("Ratings", $RatingsGroupFromLeft, $GroupsFromTop, $GroupWidth, $GroupHeight)
Global $List = GUICtrlCreateList("", $FileListFromLeft, $ListsFromTop, $ListWidth, $ListHeight)
$ListRatings = GUICtrlCreateList("", $RatingsListFromLeft, $ListsFromTop, $ListWidth, $ListHeight)

GUISetState(@SW_SHOW)

; GUI loop
While 1
    $msg = GUIGetMsg()
	_MouseSetOnEvent($MOUSE_PRIMARYDBLCLK_EVENT, '_leftClk_Event')
    Switch $msg
        Case $Checkbox_1, $Checkbox_2, $Checkbox_3
            GUICtrlSetData($List, "")
            Access($Checkbox_1)
            Access($Checkbox_2)
            Access($Checkbox_3)
			For $i=0 to UBound($List)
				If _GUICtrlListBox_GetSel($List,$i) = true Then
				$items = _GUICtrlListBox_GetSelItemsText($List)
				showMessage($items)
				Endif
			Next
        Case $GUI_EVENT_CLOSE ; Close GUI
            ExitLoop
        Case $idAddFile
            $sFileToBeCopied = FileOpenDialog("Select Files", @ScriptDir, "Text Files(*.txt)", 5)
			FileCopy($sFileToBeCopied, $pathToFiles)
            If @error Then ContinueLoop	
            $AdoCon = ObjCreate("ADODB.Connection")
            $AdoCon.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & $pathToDB)
            $AdoRs = ObjCreate("ADODB.Recordset")
            $AdoRs.CursorType = 2
            $AdoRs.LockType = 3
            $AdoRs.Open("SELECT * FROM " & $Table_Name&";", $AdoCon) 
            $aFiles = StringSplit($sFileToBeCopied, "|")
            Switch $aFiles[0]
                Case 1
                    $AdoRs.AddNew
                    $AdoRs.Fields("Feld1").value = StringTrimLeft($aFiles[1], StringInStr($aFiles[1], "\", 0, -1))
                    If BitAnd(GUICtrlRead($Checkbox_1),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld2").value = GUICtrlRead($Checkbox_1, 1)
                    If BitAnd(GUICtrlRead($Checkbox_2),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld3").value = GUICtrlRead($Checkbox_2, 1)
                    If BitAnd(GUICtrlRead($Checkbox_3),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld4").value = GUICtrlRead($Checkbox_3, 1)
                    $AdoRs.Fields("FilePath").value = $pathToFiles&"\"&StringTrimLeft($aFiles[1], StringInStr($aFiles[1], "\", 0, -1))
					$AdoRs.Update
                Case 2 To $aFiles[0]
                    For $i = 2 To $aFiles[0]
                        $AdoRs.AddNew
                        $AdoRs.Fields("Feld1").value = $aFiles[$i]
                        If BitAnd(GUICtrlRead($Checkbox_1),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld2").value = GUICtrlRead($Checkbox_1, 1)
                        If BitAnd(GUICtrlRead($Checkbox_2),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld3").value = GUICtrlRead($Checkbox_2, 1)
                        If BitAnd(GUICtrlRead($Checkbox_3),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld4").value = GUICtrlRead($Checkbox_3, 1)
                        $AdoRs.Fields("FilePath").value = $pathToFiles&"\"&StringTrimLeft($aFiles[$i], StringInStr($aFiles[$i], "\", 0, -1))
						$AdoRs.Update
                    Next
            EndSwitch
            $AdoRs.close
            $AdoCon.Close
            _DBUpdate()
            GUICtrlSetData($List, "")
            Access($Checkbox_1)
            Access($Checkbox_2)
            Access($Checkbox_3)
		case $idDownloadFile
			$testSel = _GUICtrlListBox_GetSelItemsText($List)
			$FileToBeDownloaded = _GUICtrlListBox_GetSelItems($List)
			showMessage($FileToBeDownloaded)
			FileCopy($FileToBeDownloaded, $pathToDownloads,0) 
			If @error Then 
			showMessage("File not selected!")
			Else 
			showMessage("Downloaded!!!")
			Endif
			; add here function for calculating ratings
			
    EndSwitch
WEnd

###########################################################################
;																		  ;
;''''''''''''''''''''''''''''''' FUNCTIONS ''''''''''''''''''''''''''''''';
;																		  ;
###########################################################################

Func Access($Checkbox)
    If GUICtrlRead($Checkbox) = $GUI_CHECKED Then
        Local $Chkbox_label = GUICtrlRead($Checkbox, 1)
        For $i = 0 To UBound($Files) - 1
            If $File_Attr[$i][0] = $Chkbox_label Or $File_Attr[$i][1] = $Chkbox_label Or $File_Attr[$i][2] = $Chkbox_label Then
                _GUICtrlListBox_AddString($List, $Files[$i])
            EndIF
        Next
    EndIF
EndFunc

Func _DBUpdate()
    $AdoCon = ObjCreate("ADODB.Connection")
    $AdoCon.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & $pathToDB)

    $AdoRs = ObjCreate("ADODB.Recordset")
    $AdoRs.CursorType = 1
    $AdoRs.LockType = 3
	
	; counting number of rows in  DB
    $AdoRs.Open("SELECT COUNT(*) FROM " & $Table_Name, $AdoCon)
    $dimension = $AdoRs.Fields(0).Value
	
	; redimensioning $Files to be able to get new values when adding files
    ReDim $Files[$dimension]
    ReDim $File_Attr[$dimension][3]
	
	; inserting values for checkboxes
    For $i = 0 To UBound($Files) - 1
        $AdoRs = ObjCreate("ADODB.Recordset")
        $AdoRs.CursorType = 1
        $AdoRs.LockType = 3
        $AdoRs.Open("SELECT * FROM " & $Table_Name & " WHERE ID = " & ($i + 1), $AdoCon)
        $Files[$i] = $AdoRs.Fields(1).Value
        $File_Attr[$i][0] = $AdoRs.Fields(2).Value
        $File_Attr[$i][1] = $AdoRs.Fields(3).Value
        $File_Attr[$i][2] = $AdoRs.Fields(4).Value
    Next
	
	; closing connection to DB
    $AdoRs.Close
    $AdoCon.Close

	; setting up checkboxes values
    Local $a = 0
    For $i = 0 To UBound($Files) - 1
        For $j = 0 To 2
            If $a < 3 And Not $File_Attr[$i][$j] = "" Then
                For $k = $a To 2
                    If $Attr_Name[$k] = $File_Attr[$i][$j] Then
                        ContinueLoop 2
                    EndIf
                Next
                $Attr_Name[$a] = $File_Attr[$i][$j]
                $a += 1
            EndIf
        Next
    Next
EndFunc

Func showMessage($message)
	MsgBox(0, "Error", $message)
EndFunc

Func downloadFiles($selectedFiles)
EndFunc

Func _leftClk_Event($List)
	$items = _GUICtrlListBox_GetSelItemsText($List)
	;if $items != ""
	showMessage($items)
	;else
	;showMessage("nista")
	;Endif
EndFunc