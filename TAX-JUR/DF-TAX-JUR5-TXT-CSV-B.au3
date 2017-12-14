
;******************************************************************************
; NAME: DF.TAX-JUR5 .TXT TO .CSV CONVERTER
; FILENAME: DF-TAX-JUR5-TXT-CSV.au3
;
; TYPE: AUTOIT    SCRIPT VERSION: V1.0        DATE: 10/09/2017
;******************************************************************************
; PROGRAM DESCRIPTION:
;
; Converts .TXT file to .CSV file
;
;******************************************************************************
; REPORT PARAMETERS:
;
; BEWARE: DOWNLOAD FORMATTING BETWEEN SOLAR AND ETERM ARE SUBSTANTIALLY DIFFERENT
;
; MASSLOAD DF.CUSTCON REPORT: DOWNLOADED AS .TXT FILE (CR/LF) WITH HEADINGS
;
;******************************************************************************
; SOURCEFILE DESCRIPTION:
;
; A .txt export report file from Eclipse Downloaded from SOLAR with HEADINGS
;
;******************************************************************************
; PROGRAM BEHAVIOR:
; Opens source .txt file and rewrites info to .csv
;
;
;******************************************************************************
; CHANGE NOTES:
;
; 10/09/2017  IT WORKS. but it is not the most elegant solution
;             did a hack job to do a read-ahead without reading past the
;             end of the array. Fix someday
;
;
;
;******************************************************************************
; USER DEFINED FUNCTIONS (INCLUDES):

#include <FileConstants.au3>

#include <MsgBoxConstants.au3>

#include <Array.au3>  ;Array function for File.au3

#include <File.au3>   ;File Path Functions

;******************************************************************************
; INPUT AND FILE MANIPULATION


;*********************************MAIN PROGRAM*********************************
MsgBox($MB_SYSTEMMODAL,"ECLIPSE REPORT CONVERTER", "DF.TAX-JUR5 .TXT TO .CSV CONVERSION", 3)

$sFilePath = GetFile()
;MsgBox ($MB_SYSTEMMODAL, "", "RETURNED VALUE: " & $sFilePath)

;Create Array of File Path Info
Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
Local $aPathSplit = _PathSplit($sFilePath, $sDrive, $sDir, $sFileName, $sExtension)
;_ArrayDisplay($aPathSplit, "_PathSplit of " & $sFilePath) ;Displays array contents
$nFilePath = $sDrive & $sDir & $sFileName & "-mod.csv"
$dFilePath = StringUpper($nFilePath)

MsgBox($MB_SYSTEMMODAL,"DEST FILE", $dFilePath, 3)

FileToArray($sFilePath)

;FileRdLine($sFilePath)

WriteFileLine($dFilePath)

MsgBox($MB_SYSTEMMODAL,"PROGRAM COMPLETE", "CONVERSION FINISHED", 2) ;End message, timeout 2 sec

;*********************************MAIN PROGRAM END******************************


;********************************FUNCTIONS SECTION******************************

Func GetFile()   ; Select File Dialog
    ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Select a single file of any type."

    ; Display an open dialog to select a file.
	Local $sFileOpenDialog = FileOpenDialog($sMessage, "C:\TEMP", "All (*.*)", $FD_FILEMUSTEXIST)

    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "No file was selected.")

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
;NOT USED UNLESS FILE IS TOO LARGE FOR AN ARRAY
Func FileRdLine($sFilePath)

    ; Open the file for reading and store the handle to a variable.
    Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
    If $hFileOpen = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
        Return False
    EndIf

    ; Read the fist line of the file using the handle returned by FileOpen.
    Local $sFileRead = FileReadLine($hFileOpen, 1)

    ; Close the handle returned by FileOpen.
    FileClose($hFileOpen)

    ; Display the first line of the file.
    MsgBox($MB_SYSTEMMODAL, "", "First line of the file:" & @CRLF & $sFileRead)

EndFunc   ;==>FileRdLine

;------------------------------------------------------------------------------

Func WriteFileLine($sFilePath)
    ; Create a constant variable in Local scope of the filepath that will be read/written to.
    ;Local Const $sFilePath = "C:\TEMP\AUTOIT-WRFILE.CSV"

    ; Create or open file for writing
    ;Local $hFileOpen = FileOpen($sFilePath, $FO_APPEND)
	Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
    If $hFileOpen = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the file.")
        Return False
    EndIf

	;Write Header (MULTILINE STRINGS CAN'T BE BROKEN WITH " _ " CONCATENATE INSTEAD)
	$wStr = "TAXCODE,EFFDATE,DESCRIPTION,TAX-EX-GRP,GLNO,RATE"
	;$wStr = $wStr & "xxx,xxx,xxx"

	FileWriteLine($hFileOpen, $wStr)

	;Declare serveral variables in advance (for syntax check)
	Local $vTaxcode, $vEffdate, $vDescript, $vExGrp, $vGlNo, $vRate
	Local $vTaxcode1, $vTaxcode2

	;Adjusts to not run past end of array
	$iEnd = UBound($aArray) - 1
	;MsgBox($MB_SYSTEMMODAL, "$iEnd: ", $iEnd, 3)

	For $i = 9 To $iEnd ; Skip 9 header lines in read-array
		;MsgBox($MB_SYSTEMMODAL, "", $aArray[$i]) ; Display the contents of the array.

		$j = $i + 1   ;Read ahead counter

		;Evaluate valid format and position marker
		$vTaxcode = StringMid ($aArray[$i],1 ,10)
		$vTaxcode = StringStripWS ($vTaxcode , 2) ;Remove trailing whitespace
		;$vTaxcode = StringStripWS ($vTaxcode, 1) ;Remove leading whitespace
		;$vTaxcode = StringReplace($vTaxcode, ",","")  ;Remove commas

		If $vTaxcode <> "" Then

			$vTaxcode1 = $vTaxcode

			$vEffdate = StringMid ($aArray[$i],12 ,10)
			;$vEffdate = StringStripWS ($vEffdate, 2) ;Remove trailing whitespace
			;$vEffdate = StringStripWS ($vEffdate, 1) ;Remove leading whitespace
			;$vEffdate = StringReplace($vEffdate, ",","")  ;Remove commas

			$vDescript = StringMid ($aArray[$i],23 ,30)
			$vDescript = StringStripWS ($vDescript, 2) ;Remove trailing whitespace
			;$vDescript = StringStripWS ($vDescript, 1) ;Remove leading whitespace
			$vDescript = StringReplace($vDescript, ",","")  ;Remove commas

			$vExGrp = StringMid ($aArray[$i],54 ,10)
			$vExGrp = StringStripWS ($vExGrp, 2) ;Remove trailing whitespace
			$vExGrp = StringStripWS ($vExGrp, 1) ;Remove leading whitespace
			;$vExGrp = StringReplace($vExGrp, ",","")  ;Remove commas

			$vGlNo = StringMid ($aArray[$i],65 ,5)
			;$vGlNo = StringStripWS ($vGlNo, 2) ;Remove trailing whitespace
			$vGlNo = StringStripWS ($vGlNo, 1) ;Remove leading whitespace
			;$vGlNo = StringReplace($vGlNo, ",","")  ;Remove commas

			$vRate = StringMid ($aArray[$i],71 ,6)
			$vRate = StringStripWS ($vRate, 2) ;Remove trailing whitespace
			;$vRate = StringStripWS ($vRate, 1) ;Remove leading whitespace
			;$vRate = StringReplace($vRate, ",","")  ;Remove commas

			;Read Taxcode of next line down to see if it is a new record or not
			;Set to "END" if last line of file
			;Added If $iEnd to fix read past end of array error
			If $i = $iEnd  Then
				$vTaxcode2 = "END"
				 ;MsgBox($MB_SYSTEMMODAL, "", $vTaxcode1, 3)
			Else
				$vTaxcode2 = StringMid ($aArray[$j],1 ,10)
				$vTaxcode2 = StringStripWS ($vTaxcode2, 2) ;Remove trailing whitespace
				;$vTaxcode2 = StringStripWS ($vTaxcode2, 1) ;Remove leading whitespace
				;$vTaxcode2 = StringReplace($vTaxcode2, ",","")  ;Remove commas
			EndIf

			;Check if next line is new record and write file if it is
			If $vTaxcode2 <> "" Then

				;Write to result file in comma separated format
				$wStr = $vTaxcode1 & "," & $vEffdate & "," & $vDescript & "," & $vExGrp _
				& "," & $vGlNo & "," & $vRate

				;Write to result file in TAB separated format
				;wStr = $aStr & @TAB  ;TAB Separated

				FileWriteLine($hFileOpen, $wStr)
			EndIf


		Else   ;Next lines of record with potential multiple rates

			If $i = $iEnd  Then
				$vTaxcode2 = "END"
				 ;MsgBox($MB_SYSTEMMODAL, "", $vTaxcode1, 3)
			Else
				$vTaxcode2 = StringMid ($aArray[$j],1 ,10)
				$vTaxcode2 = StringStripWS ($vTaxcode2, 2) ;Remove trailing whitespace
				;$vTaxcode2 = StringStripWS ($vTaxcode2, 1) ;Remove leading whitespace
				;$vTaxcode2 = StringReplace($vTaxcode2, ",","")  ;Remove commas
			EndIf

			;Check if next line is new record
			If $vTaxcode2 <> "" Then

				$vExGrp = StringMid ($aArray[$i],54 ,10)
				$vExGrp = StringStripWS ($vExGrp, 2) ;Remove trailing whitespace
				$vExGrp = StringStripWS ($vExGrp, 1) ;Remove leading whitespace
				;$vExGrp = StringReplace($vExGrp, ",","")  ;Remove commas

				$vGlNo = StringMid ($aArray[$i],65 ,5)
				;$vGlNo = StringStripWS ($vGlNo, 2) ;Remove trailing whitespace
				$vGlNo = StringStripWS ($vGlNo, 1) ;Remove leading whitespace
				;$vGlNo = StringReplace($vGlNo, ",","")  ;Remove commas

				$vRate = StringMid ($aArray[$i],71 ,6)
				$vRate = StringStripWS ($vRate, 2) ;Remove trailing whitespace
				;$vRate = StringStripWS ($vRate, 1) ;Remove leading whitespace
				;$vRate = StringReplace($vRate, ",","")  ;Remove commas

				;Write to result file in comma separated format
				$wStr = $vTaxcode1 & "," & $vEffdate & "," & $vDescript & "," & $vExGrp _
				& "," & $vGlNo & "," & $vRate

				;Write to result file in TAB separated format
				;wStr = $aStr & @TAB  ;TAB Separated

				FileWriteLine($hFileOpen, $wStr)
			EndIf


		EndIf  ;$vTaxcode <> ""

	Next ;For $i


    ; Close the handle returned by FileOpen.
    FileClose($hFileOpen)

    ; Display the contents of the file passing the filepath to FileRead instead of a handle returned by FileOpen.
    ;MsgBox($MB_SYSTEMMODAL, "", "Contents of the file:" & @CRLF & FileRead($sFilePath))

    ; Delete the temporary file.
    ;FileDelete($sFilePath)
EndFunc   ;==>Example



;******************************END OF FUNCTIONS SECTION************************
