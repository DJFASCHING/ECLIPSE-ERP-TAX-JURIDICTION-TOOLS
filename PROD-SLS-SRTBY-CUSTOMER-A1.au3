;******************************************************************************
; NAME: ECLIPSE PRODUCT SALES REPORT CONVERTER
; FILENAME: PROD-SLS-SRTBY-PRODUCT-XX.AU3
; REVISION: 1.0
;
; TYPE: AUTOIT    VERSION:   V 3.3.12.0        DATE: 2018:02:22
;******************************************************************************
; PROGRAM DESCRIPTION:
;
; Reads .TXT format file download from Eclipse Hold File
;
; This version reads source file directly into an array
;
;******************************************************************************
; REPORT PARAMETERS:
;
;
;******************************************************************************
; SOURCEFILE DESCRIPTION:
;
; Eclipse Report Downloaded in ASCII - LF = CR/LF  text format with headings
;
;******************************************************************************
; PROGRAM BEHAVIOR:
; Reads text file into array, cleans  up data and writes to TAB format file
;
;
;******************************************************************************
; CHANGE NOTES:
;
; 2018:02:22 CREATED
; 2018:02:22 IT WORKS!
;
;******************************************************************************
; TO-DO NOTES:
;
;
;******************************************************************************
; PRODUCT SALES REPORT PARAMETERS:
;
; SELECT BY: 				PRODUCT
; SORT BY: 					CUSTOMER
; SELECT BRANCH: 			PRICING
; DETAIL/SUMMARY: 			DETAIL
; STATUS: 					ALL
; SERIAL NUMBERS: 			NONE
; SHOW COSTS: 				YES
; SHOW KITS AS COMPONENTS: 	NO
;
; With ths SORT-BY as Customer the script has to fill in the missing succeeding
; CUSTOMER ID and CUSTOMER NAME to product lines
;
;***********************************************************************************
;------------------------------------------------------------------------------------
;COL	EXCEL	FIELDNAME							VARNAME             POS		LEN
;------------------------------------------------------------------------------------
;001	A		CUST #								$sA1				1		6
;002	B		CUSTOMER NAME						$sB1				9		26
;003	C		INVOICE #							$sC1				37		14
;004	D		WHS									$sD1				53		4
;005	E		SHIPDATE							$sE1				58		9
;006	F		PROD NO								$sF1				70		7
;007	G		PRODUCT								$sG1				79		31
;008	H		QTY SHIPPED							$sH1				114		9
;009	I		EXT AMOUNT							$sI1				126		16
;010	J		OVERRIDE FLAG						$sJ1				142		1
;011	K		EXT COST							$sK1				144		16
;012	L		GP$									$sL1				162		13
;013	M		GP%									$sM1				177		6
;***********************************************************************************
;***********************************************************************************
; USER DEFINED FUNCTIONS (INCLUDES):

;FOR USER INTERACTION:
#include <MsgBoxConstants.au3>
;FOR GUI:
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
;FOR SQLITE DB:
;#include <SQLite.au3>
;#include <SQLite.dll.au3>
;FOR FILE OPERATIONS:
#include <File.au3>  ;File O/R/W Ops
#include <FileConstants.au3>  ;File Data
;FOR ARRAY / FILE EXPORT FUNCTIONS:
#include <Array.au3>


;******************************************************************************
; INITIALIZATION:
;******************************************************************************

;GLOBALS:
Global $LogPath = @ScriptDir & '\CMDLOG.TXT'  ;gets rewritten later in script

;LOCALS:
Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = "",  $iFileExists = ""

;KILLSWITCH:
HotKeySet('{NUMPADSUB}', 'Hotkey1')   ;Number Pad Subtract Key
;MsgBox(0, 'KILLSWITCH', 'TERMINATE WITH NUMPAD - ',2)

;*********************************MAIN PROGRAM*********************************
#Region ### START Koda GUI section ### Form=E:\DATA - Copy\_SCRIPTS\AUTOIT\KODA-FORMS\FORM-FILECONV-A3.kxf
$MAIN = GUICreate("MAIN", 580, 400, 278, 148)

$mLbTitle = GUICtrlCreateLabel("ECLIPSE PRODUCT SALES REPORT CONVERSION", 24, 16, 300, 17)
GUICtrlSetFont(-1, 8, 800, 4, "MS Sans Serif")

$mLbDesc = GUICtrlCreateLabel("CONVERT XXX.TXT TO .TAB FORMAT FOR EXCEL IMPORT", 24, 48, 400, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

$mBSelect = GUICtrlCreateButton("SELECT", 48, 88, 75, 25)

$mBConvert = GUICtrlCreateButton("CONVERT", 232, 88, 75, 25)

$mBLog = GUICtrlCreateButton("LOG", 424, 88, 75, 25)

$mLbSource = GUICtrlCreateLabel("SOURCE FILE", 40, 144, 85, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mISource = GUICtrlCreateInput("SOURCE FILE", 133, 144, 425, 21)

$mLbConvert = GUICtrlCreateLabel("CONVERTED FILE", 16, 192, 110, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mIConvert = GUICtrlCreateInput("CONVERTED FILE", 133, 190, 425, 21)

$mLbLog = GUICtrlCreateLabel("ERROR LOG", 48, 240, 77, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mILog = GUICtrlCreateInput("LOG FILE", 133, 238, 425, 21)

$mLbLines = GUICtrlCreateLabel("LINES / PROC", 32, 288, 88, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mILines = GUICtrlCreateInput("LINES", 133, 288, 105, 21)

$mLbStat = GUICtrlCreateLabel("STATUS", 256, 288, 53, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mIStat = GUICtrlCreateInput("STATUS", 316, 287, 177, 21)

$mLbNote = GUICtrlCreateLabel("NOTE", 80, 336, 38, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mINote = GUICtrlCreateInput("NOTE", 133, 335, 361, 21)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


GUICtrlSetData($mINote, "[NUMPAD]  -  TO TERMINATE")  ;Display note in GUI

While 1
	$nMsg = GUIGetMsg()  ;Idles CPU when there is no events waiting
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $mBSelect
			$sFilePath = GetFile()
			Local $iFileExists = FileExists($sFilePath)
			If $iFileExists Then
				Local $aPathSplit = _PathSplit($sFilePath, $sDrive, $sDir, $sFileName, $sExtension)
				;_ArrayDisplay($aPathSplit, "_PathSplit of " & $sFilePath) ;Displays array contents
				Local $nFilePath = $sDrive & $sDir & $sFileName & "-tab.txt"
				Local $dFilePath = StringUpper($nFilePath)
				;LogFile Creation / Open
				$LogPath = $sDrive & $sDir & $sFileName & "-log.txt"  ;Modify Log file path
				$LogPath = StringUpper($LogPath)
				;Global $hLogOpen = FileOpen($LogPath, $FO_OVERWRITE)
				;If $hLogOpen = -1 Then
				;	MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the LOG file.")
				;EndIf
				GUICtrlSetData($mISource, $sFilePath)  ;GUI Info
				GUICtrlSetData($mIConvert, $dFilePath)  ;GUI Info
				GUICtrlSetData($mILog, $LogPath)  ;GUI Info
			EndIf
		Case $mBConvert
			If $iFileExists Then
				FileToArray($sFilePath)
				RWFile($sFilePath,$dFilePath)
			Else
				MsgBox($MB_SYSTEMMODAL, "", "SELECT SOURCE FILE FIRST", 2)
			EndIf
		Case $mBLog
			ShellExecute("Notepad.exe", $LogPath)

	EndSwitch
WEnd

;*********************************MAIN PROGRAM END******************************


;********************************FUNCTIONS SECTION******************************
;KILLSWITCH:

Func Hotkey1()
     MsgBox(0, 'EXIT', 'PROGRAM TERMINATED',2)
	 ;Close Database
	 ;_SQLite_Close($hDskDb)  ;Close Opened DB file
	; Close the opened Source Data File.
    ;FileClose($hFileOpen)
	 Exit
EndFunc     ;==>Hotkey1

;------------------------------------------------------------------------------

Func GetFile()   ; Select File Dialog
    ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Select a single file of any type."

    ; Display an open dialog to select a file.
	Local $sFileOpenDialog = FileOpenDialog($sMessage, "C:\TEMP", "All (*.*)", $FD_FILEMUSTEXIST)

    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "No file was selected.")
		Exit

        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
    Else
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)

        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
        $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the selected file.
        ;MsgBox($MB_SYSTEMMODAL, "", "You chose the following file:" & @CRLF & $sFileOpenDialog)
    EndIf
	Return $sFileOpenDialog ;Return Value
EndFunc   ;==>GetFile

;------------------------------------------------------------------------------

Func FileToArray($sFilePath)
    ; Read the current script file into an array using the filepath.
    Global $aArray = FileReadToArray($sFilePath)
    If @error Then
        MsgBox($MB_SYSTEMMODAL, "", "There was an error reading the file. @error: " & @error) ; An error occurred reading the current script file.

;	Else
;        For $i = 0 To UBound($aArray) - 1 ; Loop through the array.
;            MsgBox($MB_SYSTEMMODAL, "", $aArray[$i]) ; Display the contents of the array.
;        Next
    EndIf
EndFunc   ;==>FileToArray


;------------------------------------------------------------------------------
;NOT USED UNLESS FILE IS TOO LARGE FOR AN ARRAY (REFERENCE ONLY)
Func FileRdLine($sFilePath)

    ; Open the file for reading and store the handle to a variable.
    Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
    If $hFileOpen = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
        Return False
    EndIf

    ; Read the fist line of the file using the handle returned by FileOpen.
    Local $sFileRead = FileReadLine($hFileOpen, 2) ;Read line 2 of file (L1 is header)

    ; Close the handle returned by FileOpen.
    FileClose($hFileOpen)

    ; Display the first line of the file.
    MsgBox($MB_SYSTEMMODAL, "", "First line of the file:" & @CRLF & $sFileRead)
	Return
EndFunc   ;==>FileRdLine

;------------------------------------------------------------------------------
;Read Source File Array and Write to Result file
Func RWFile($sFilePath, $dFilePath)

	;Set Progress Counters
	$iCtr1 = 0
	$iCtr2 = 0
	GUICtrlSetData($mILines, $iCtr2)  ;Set Initial Progress in GUI
	GUICtrlSetData($mIStat, "RUNNING")  ;Set Initial Status in GUI

	;Open Error log file
	Local $hLogOpen = FileOpen($LogPath, $FO_OVERWRITE)
	If $hLogOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the LOG file.")
	EndIf

	FileWriteLine($hLogOpen, "ERROR LOG CONTENTS: ")

    ; Create or open file for writing
    ;Local $hFileOpenW = FileOpen($sFilePath, $FO_APPEND)
	Local $hFileOpenW = FileOpen($dFilePath, $FO_OVERWRITE)
    If $hFileOpenW = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the file.")
        Return False
    EndIf


	;Write Header
	$sWrt = "CUSTNO" & @TAB
	$sWrt = $sWrt & "CUSTOMER" & @TAB
	$sWrt = $sWrt & "INVNO" & @TAB
	$sWrt = $sWrt & "WHS" & @TAB
	$sWrt = $sWrt & "SHIPDATE" & @TAB
	$sWrt = $sWrt & "ECLID" & @TAB
	$sWrt = $sWrt & "PRODUCT" & @TAB
	$sWrt = $sWrt & "QTY" & @TAB
	$sWrt = $sWrt & "EXT-AMT" & @TAB
	$sWrt = $sWrt & "PRC-OVRD" & @TAB
	$sWrt = $sWrt & "EXT-COST" & @TAB
	$sWrt = $sWrt & "GP-DOL" & @TAB
	$sWrt = $sWrt & "GP-PCT"

	;Write Header to file
	FileWriteLine($hFileOpenW, $sWrt)

	;Declare serveral variables in advance (for syntax check)
	Local $sMrk1, $sMrk2, $sA1, $sB1, $sC1, $sD1, $sE1, $sF1, $sG1, $sH1, $sI1, $sJ1, $sK1, $sL1, $sM1


	For $i = 6 To UBound($aArray) - 1 ; Skip 7 header lines in read-array
		;MsgBox($MB_SYSTEMMODAL, "", $aArray[$i]) ; Display the contents of the array.

		;Evaluate valid format and position marker
		$sMrk1 = StringMid ($aArray[$i],1 ,1) ; & mark at front of Customer ID
		;$sMrk1 = StringStripWS ($sMrk1 , 2) ;Remove trailing whitespace
		;$sMrk1 = StringStripWS ($sMrk1, 1) ;Remove leading whitespace
		;$sMrk1 = StringReplace($sMrk1, ",","")  ;Remove commas

				;Evaluate valid format and position marker
		$sMrk2 = StringMid ($aArray[$i],69 ,1) ; ^ mark at front of Prod No
		;$sMrk2 = StringStripWS ($sMrk2 , 2) ;Remove trailing whitespace
		;$sMrk2 = StringStripWS ($sMrk2, 1) ;Remove leading whitespace
		;$sMrk2 = StringReplace($sMrk2, ",","")  ;Remove commas

		If $sMrk1 == "&" Then

			;001	A		CUST #								$sA1				2		5
			$sA1 = StringMid ($aArray[$i],2 ,5)
			$sA1 = StringStripWS ($sA1, 2) ;Remove trailing whitespace
			;$sA1 = StringStripWS ($sA1, 1) ;Remove leading whitespace
			;$sA1 = StringReplace($sA1, ",","")  ;Remove commas

			;002	B		CUSTOMER NAME						$sB1				9		26
			$sB1 = StringMid ($aArray[$i],9 ,26)
			$sB1 = StringStripWS ($sB1, 2) ;Remove trailing whitespace
			;$sB1 = StringStripWS ($sB1, 1) ;Remove leading whitespace
			$sB1 = StringReplace($sB1, ",","")  ;Remove commas

			;003	C		INVOICE #							$sC1				37		14
			$sC1 = StringMid ($aArray[$i],37 ,14)  ;
			$sC1 = StringStripWS ($sC1, 2) ;Remove trailing whitespace
			;$sC1 = StringStripWS ($sC1, 1) ;Remove leading whitespace
			;$sC1 = StringReplace($sC1, ",","")  ;Remove commas

			;004	D		WHS									$sD1				53		4
			$sD1 = StringMid ($aArray[$i],53 ,4)
			;$sD1 = StringStripWS ($sD1, 2) ;Remove trailing whitespace
			$sD1 = StringStripWS ($sD1, 1) ;Remove leading whitespace
			;$sD1 = StringReplace($sD1, ",","")  ;Remove commas

			;005	E		SHIPDATE							$sE1				58		9
			$sE1 = StringMid ($aArray[$i],58 ,9)
			;$sE1 = StringStripWS ($sE1, 2) ;Remove trailing whitespace
			$sE1 = StringStripWS ($sE1, 1) ;Remove leading whitespace
			;$sE1 = StringReplace($$sE1, ",","")  ;Remove commas

			;006	F		PROD NO								$sF1				70		7
			$sF1 = StringMid ($aArray[$i],68 ,9)
			$sF1 = StringStripWS ($sF1, 2) ;Remove trailing whitespace
			$sF1 = StringStripWS ($sF1, 1) ;Remove leading whitespace
			$sF1 = StringReplace($sF1, "^","")  ;Remove ^ Carrot

			;007	G		PRODUCT								$sG1				79		31
			$sG1 = StringMid ($aArray[$i],79 ,31)
			$sG1 = StringStripWS ($sG1, 2) ;Remove trailing whitespace
			;$sG1 = StringStripWS ($sG1, 1) ;Remove leading whitespace
			$sG1 = StringReplace($sG1, ",","")  ;Remove , Commas

			;008	H		QTY SHIPPED							$sH1				114		9
			$sH1 = StringMid ($aArray[$i],114 ,9)
			;$sH1 = StringStripWS ($sH1, 2) ;Remove trailing whitespace
			$sH1 = StringStripWS ($sH1, 1) ;Remove leading whitespace
			$sH1 = StringReplace($sH1, ",","")  ;Remove , Commas

			;009	I		EXT AMOUNT							$sI1				126		16
			$sI1 = StringMid ($aArray[$i],126 ,16)
			;$sI1 = StringStripWS ($sI1, 2) ;Remove trailing whitespace
			$sI1 = StringStripWS ($sI1, 1) ;Remove leading whitespace
			$sI1 = StringReplace($sI1, ",","")  ;Remove , Commas

			;010	J		OVERRIDE FLAG						$sJ1				142		1
			$sJ1 = StringMid ($aArray[$i],142 ,1)
			;$sJ1 = StringStripWS ($sJ1, 2) ;Remove trailing whitespace
			;$sJ1 = StringStripWS ($sJ1, 1) ;Remove leading whitespace
			;$sJ1 = StringReplace($sJ1, ",","")  ;Remove commas
			If $sJ1 == "*" Then  ;Change Price Overide Indicator
				$sJ1 = "Y"
			Else
				$sJ1 = "N"
			EndIf

			;011	K		EXT COST							$sK1				144		16
			$sK1 = StringMid ($aArray[$i],144 ,16)
			;$sK1 = StringStripWS ($sK1, 2) ;Remove trailing whitespace
			$sK1 = StringStripWS ($sK1, 1) ;Remove leading whitespace
			$sK1 = StringReplace($sK1, ",","")  ;Remove , Commas

			;012	L		GP$									$sL1				162		13
			$sL1 = StringMid ($aArray[$i],162 ,13)
			;$sL1 = StringStripWS ($sL1, 2) ;Remove trailing whitespace
			$sL1 = StringStripWS ($sL1, 1) ;Remove leading whitespace
			$sL1 = StringReplace($sL1, ",","")  ;Remove , Commas

			;013	M		GP%									$sM1				177		7
			$sM1 = StringMid ($aArray[$i],177 ,7)
			;$sM1 = StringStripWS ($sM1, 2) ;Remove trailing whitespace
			$sM1 = StringStripWS ($sM1, 1) ;Remove leading whitespace
			;$sM1 = StringReplace($sM1, ",","")  ;Remove , Commas


			;Write Line to Result File
			$sWrt = $sA1 & @TAB
			$sWrt = $sWrt & $sB1 & @TAB
			$sWrt = $sWrt & $sC1 & @TAB
			$sWrt = $sWrt & $sD1 & @TAB
			$sWrt = $sWrt & $sE1 & @TAB
			$sWrt = $sWrt & $sF1 & @TAB
			$sWrt = $sWrt & $sG1 & @TAB
			$sWrt = $sWrt & $sH1 & @TAB
			$sWrt = $sWrt & $sI1 & @TAB
			$sWrt = $sWrt & $sJ1 & @TAB
			$sWrt = $sWrt & $sK1 & @TAB
			$sWrt = $sWrt & $sL1 & @TAB
			$sWrt = $sWrt & $sM1

			FileWriteLine($hFileOpenW, $sWrt)


		ElseIf $sMrk2 == "^" Then

			;003	C		INVOICE #							$sC1				37		14
			$sC1 = StringMid ($aArray[$i],37 ,14)  ;
			$sC1 = StringStripWS ($sC1, 2) ;Remove trailing whitespace
			;$sC1 = StringStripWS ($sC1, 1) ;Remove leading whitespace
			;$sC1 = StringReplace($sC1, ",","")  ;Remove commas

			;004	D		WHS									$sD1				53		4
			$sD1 = StringMid ($aArray[$i],53 ,4)
			;$sD1 = StringStripWS ($sD1, 2) ;Remove trailing whitespace
			$sD1 = StringStripWS ($sD1, 1) ;Remove leading whitespace
			;$sD1 = StringReplace($sD1, ",","")  ;Remove commas

			;005	E		SHIPDATE							$sE1				58		9
			$sE1 = StringMid ($aArray[$i],58 ,9)
			;$sE1 = StringStripWS ($sE1, 2) ;Remove trailing whitespace
			$sE1 = StringStripWS ($sE1, 1) ;Remove leading whitespace
			;$sE1 = StringReplace($$sE1, ",","")  ;Remove commas

			;006	F		PROD NO								$sF1				70		7
			$sF1 = StringMid ($aArray[$i],68 ,9)
			$sF1 = StringStripWS ($sF1, 2) ;Remove trailing whitespace
			$sF1 = StringStripWS ($sF1, 1) ;Remove leading whitespace
			$sF1 = StringReplace($sF1, "^","")  ;Remove ^ Carrot

			;007	G		PRODUCT								$sG1				79		31
			$sG1 = StringMid ($aArray[$i],79 ,31)
			$sG1 = StringStripWS ($sG1, 2) ;Remove trailing whitespace
			;$sG1 = StringStripWS ($sG1, 1) ;Remove leading whitespace
			$sG1 = StringReplace($sG1, ",","")  ;Remove , Commas

			;008	H		QTY SHIPPED							$sH1				114		9
			$sH1 = StringMid ($aArray[$i],114 ,9)
			;$sH1 = StringStripWS ($sH1, 2) ;Remove trailing whitespace
			$sH1 = StringStripWS ($sH1, 1) ;Remove leading whitespace
			$sH1 = StringReplace($sH1, ",","")  ;Remove , Commas

			;009	I		EXT AMOUNT							$sI1				126		16
			$sI1 = StringMid ($aArray[$i],126 ,16)
			;$sI1 = StringStripWS ($sI1, 2) ;Remove trailing whitespace
			$sI1 = StringStripWS ($sI1, 1) ;Remove leading whitespace
			$sI1 = StringReplace($sI1, ",","")  ;Remove , Commas

			;010	J		OVERRIDE FLAG						$sJ1				142		1
			$sJ1 = StringMid ($aArray[$i],142 ,1)
			;$sJ1 = StringStripWS ($sJ1, 2) ;Remove trailing whitespace
			;$sJ1 = StringStripWS ($sJ1, 1) ;Remove leading whitespace
			;$sJ1 = StringReplace($sJ1, ",","")  ;Remove commas
			If $sJ1 == "*" Then  ;Change Price Overide Indicator
				$sJ1 = "Y"
			Else
				$sJ1 = "N"
			EndIf

			;011	K		EXT COST							$sK1				144		16
			$sK1 = StringMid ($aArray[$i],144 ,16)
			;$sK1 = StringStripWS ($sK1, 2) ;Remove trailing whitespace
			$sK1 = StringStripWS ($sK1, 1) ;Remove leading whitespace
			$sK1 = StringReplace($sK1, ",","")  ;Remove , Commas

			;012	L		GP$									$sL1				162		13
			$sL1 = StringMid ($aArray[$i],162 ,13)
			;$sL1 = StringStripWS ($sL1, 2) ;Remove trailing whitespace
			$sL1 = StringStripWS ($sL1, 1) ;Remove leading whitespace
			$sL1 = StringReplace($sL1, ",","")  ;Remove , Commas

			;013	M		GP%									$sM1				177		7
			$sM1 = StringMid ($aArray[$i],177 ,7)
			;$sM1 = StringStripWS ($sM1, 2) ;Remove trailing whitespace
			$sM1 = StringStripWS ($sM1, 1) ;Remove leading whitespace
			;$sM1 = StringReplace($sM1, ",","")  ;Remove , Commas


			;Write Line to Result File
			$sWrt = $sA1 & @TAB
			$sWrt = $sWrt & $sB1 & @TAB
			$sWrt = $sWrt & $sC1 & @TAB
			$sWrt = $sWrt & $sD1 & @TAB
			$sWrt = $sWrt & $sE1 & @TAB
			$sWrt = $sWrt & $sF1 & @TAB
			$sWrt = $sWrt & $sG1 & @TAB
			$sWrt = $sWrt & $sH1 & @TAB
			$sWrt = $sWrt & $sI1 & @TAB
			$sWrt = $sWrt & $sJ1 & @TAB
			$sWrt = $sWrt & $sK1 & @TAB
			$sWrt = $sWrt & $sL1 & @TAB
			$sWrt = $sWrt & $sM1

			FileWriteLine($hFileOpenW, $sWrt)


		Else
			;Skipped lines written to log
			FileWriteLine($hLogOpen, $aArray[$i])

		EndIf  ;End If $sMrk1 == "&"


		;Update Progress counter
		$iCtr1 = $iCtr1 + 1

		;Progress counter
		If $iCtr1 == 50 Then
			$iCtr2 = $iCtr2 + $iCtr1
			GUICtrlSetData($mILines, $iCtr2)  ;Display Progress in GUI
			$iCtr1 = 0
		EndIf

	Next ;For $i


	;Final counter update
	$iCtr2 = $iCtr2 + $iCtr1
	GUICtrlSetData($mILines, $iCtr2)  ;Display Progress in GUI
	GUICtrlSetData($mIStat, "DONE")  ;Set Status in GUI

    ; Read the fist line of the file using the handle returned by FileOpen.
    ;Local $sFileRead = FileReadLine($hFileOpenR, 2) ;Read line 2 of file (L1 is header)

    ; Display the first line of the file.
    ;MsgBox($MB_SYSTEMMODAL, "", "First line of the file:" & @CRLF & $sFileRead)

	; Close the handle returned by FileOpen.
	;Close Destination Write File
	FileClose($hFileOpenW)
	;Close the log file
	FileClose($hLogOpen)

	Return
EndFunc   ;==>FileRdLine

;------------------------------------------------------------------------------



