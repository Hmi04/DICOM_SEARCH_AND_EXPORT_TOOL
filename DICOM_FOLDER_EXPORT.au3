#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GuiListView.au3>
#include <GuiComboBox.au3>
#include <File.au3>
#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>

#include <GuiListBox.au3>
#include <EditConstants.au3>
#include <Misc.au3>

;Opt("GUIOnEventMode", 1) ;too much trouble

Global $hGui = GUICreate("DICOM SEARCH & USB SAVE TOOL", 500, 500, 0, 0, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX, $WS_MAXIMIZEBOX), 0x00000010)
Global $sSelectedFile=""
Global $sFolder=""
Global $DoubleClicked   = False
Global $sFolder = "C:\DICOM"
Global $aOldFolderArrayOnly=""

GUISetState()
$lv = GUICtrlCreateListView("", 2, 0, 496,500-50)

_GUICtrlListView_AddColumn($lv, "Name", 140)
_GUICtrlListView_AddColumn($lv, "Gender", 20)
_GUICtrlListView_AddColumn($lv, "Birthday", 140)
_GUICtrlListView_AddColumn($lv, "ID", 140)
_GUICtrlListView_AddColumn($lv, "Accession", 140)
_GUICtrlListView_AddColumn($lv, "Doctor", 300)
_GUICtrlListView_AddColumn($lv, "Folder", 300)

GUIRegisterMsg( $WM_NOTIFY, "WM_NOTIFY" )

;GUICtrlCreateLabel("Patient Name", 2, 248, -1, 14)

$input1 = GUICtrlCreateInput("", 2, 500-44, 500-100-2, 20)
$button1 = GUICtrlCreateButton("Search", 500-100-1, 500-44-1,100,22,$BS_DEFPUSHBUTTON)
;GUICtrlSetOnEvent(-1, "_Search")
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)


Local $aDetectedDriveLetters = DriveGetDrive($DT_ALL)
If @error Then
ConsoleWriteError("error")
Else
	$i=$aDetectedDriveLetters[0]
EndIf

$input2 = GUICtrlCreateInput(StringUpper($aDetectedDriveLetters[$i])&"\", 2, 500-22, 500-380, 20)
$button2 = GUICtrlCreateButton("...", 500-380+2, 500-22-1,40,22)
;;GUICtrlSetOnEvent(-1, "_ChangeSaveFolder")

$button3 = GUICtrlCreateButton("Folder Name to Clipboard", 500-380+40+2, 500-22-1,200,22)
$button4 = GUICtrlCreateButton("Copy to USB", 500-380+40+200+2-1, 500-22-1,138,22)
;GUICtrlSetOnEvent(-1, "_CopytoUSB")

If UBound($CmdLine)>1 Then
$sFolder = $CmdLine[1]
Else
$sTempFolder=$sFolder
$sFolder=FileSelectFolder("Choose Save/Export Folder",$aDetectedDriveLetters[$i])

;invalid folder path
if(StringLen($sFolder)<2) Then
   $sFolder=$sTempFolder
EndIf

;remove extra slash from path
if(StringRight($sFolder, 1) <> "\") Then
$sFolder=$sFolder&"\"
EndIf

EndIf

loadlist()

 While 1
      Switch GUIGetMsg()
		  Case $GUI_EVENT_CLOSE
                ExitLoop
 		  Case $button1
 			;ConsoleWrite("button1")
			_Search()
 		  Case $button2
 			;ConsoleWrite("button2")
			_ChangeSaveFolder()
 		  Case $button3
			_OpenListviewFolder()
			if($sSelectedFile>0) then
			ClipPut($sFolder & $sSelectedFile) ;& "\"
			EndIf
 		  Case $button4
 			;ConsoleWrite("button4")
			$sSaveFolder=GUICtrlRead($input2)
			_CopyDirWithProgress($sFolder&"\"&$sSelectedFile,$sSaveFolder&"\"&$sSelectedFile)

      EndSwitch

    If $DoubleClicked Then
        _OpenListviewFolder()
        $DoubleClicked = False
		ShellExecute($sFolder & "\" & $sSelectedFile)
	 EndIf


   ;; If _IsPressed("{Enter}") Then _Search()


 WEnd
 GUIDelete();
Exit



Func _ChangeSaveFolder()
$sSaveFolder=GUICtrlRead($input2)
$sSaveFolder=FileSelectFolder("Choose Save/Export Folder",$sSaveFolder)

;invalid folder path
if(StringLen($sSaveFolder)<2) Then
   $sSaveFolder=GUICtrlRead($input2)
EndIf

;remove extra slash from path
if(StringRight($sSaveFolder, 1) <> "\") Then
$sSaveFolder=$sSaveFolder&"\"
EndIf

ControlSetText("","",$input2,$sSaveFolder)


EndFunc

Func Filter($c1)
     loadlist()
     $itemcount = _GUICtrlListView_GetItemCount($lv)
     For $z = $itemcount - 1 To 0 Step -1
		  $item = _GUICtrlListView_GetItemTextString($lv, $z) ;searches whole listbox
          ;$item = _GUICtrlListView_GetItemText($lv, $z) ;searches just the first column
;          If ($c1 = "" Or $c1 = $item or StringInStr($item,$c1)) Then
          If ($c1 = "" Or StringInStr($item,$c1)) Then
                    ContinueLoop
          Else
               _GUICtrlListView_DeleteItem($lv, $z)
          EndIf
     Next
  EndFunc   ;==>Filter


Func _Search()
;loadlist()
Filter(GUICtrlRead($input1))
EndFunc

Func loadlist()
;local $aFolderArray = _FileListToArrayRec($sFolder, "*_digital_look.dcm", $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT) ;vidar only
local $aFolderArray = _FileListToArrayRec($sFolder, "*.dcm", $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT) ;all DICOM files
;Local $aFilenameArray

;_SendMessage($lv,$WM_SETREDRAW,0,0)
;GUICtrlSetData($lv, "Hide")
     _GUICtrlListView_DeleteAllItems($lv)


For $i = 1 To UBound($aFolderArray) - 1
 Local $aFolderArrayOnly= StringRegExp(String($aFolderArray[$i]), '[^\\]*', $STR_REGEXPARRAYMATCH)
 ;$aFilenameArray = StringRegExp(String($aFolderArray[$i]), '(?i)(\x5C)(.*?)(.dcm)', $STR_REGEXPARRAYMATCH)
if UBound($aFolderArrayOnly) >0 Then
   if $aFolderArrayOnly[0]<>$aOldFolderArrayOnly then
	  GUICtrlCreateListViewItem(ReadHeaderDICOM($sFolder &"\"& $aFolderArray[$i])&"|"&$aFolderArrayOnly[0], $lv)
	  ;GUICtrlSetOnEvent(-1, "_OpenListviewFolder")
	  $aOldFolderArrayOnly=$aFolderArrayOnly[0]
;ConsoleWrite($aFolderArrayOnly[0])
   EndIf
EndIf

Next
;_SendMessage($lv,$WM_SETREDRAW,1,0)



EndFunc


Func _OpenListviewFolder()
$sSelectedFile=StringRegExp(_GUICtrlListView_GetItemTextString($lv), '([^|]+$)', $STR_REGEXPARRAYMATCH)
If UBound($sSelectedFile)>0 Then
$sSelectedFile=$sSelectedFile[0]
EndIf

;   ConsoleWrite($sSelectedFile[0])
EndFunc

Func ReadHeaderDICOM($filename)
Local $offset=""
Local $sDoctorName,$sPatientName,$sID,$sDOB,$sSex,$sAccession=""
Local $in=FileOpen($filename, 16) ; 16+0=Read binary
;Local $data = FileRead($in,0x7E0);ONLY READ HEADER, WHOLE FILE IS 50MB!
Local $DICOMHEADER = FileRead($in,0x560);ONLY READ HEADER, WHOLE FILE IS 50MB!

FileClose($in)


;Local $aArray = StringRegExp(String($data), '(08009000)(.*?)18001610', $STR_REGEXPARRAYMATCH) ;parse whole header from physician to vendor
;Local $aArray = StringRegExp(String($data), '(02000200)(.*?)28000400', $STR_REGEXPARRAYMATCH) ;parse whole header for more common tags
;if UBound($aArray)<1 then
;   Return $sPatientName&"|"&string($sSex)&"|"&$sDOB&"|"&$sDoctorName ;exit because header wasn't found
;EndIf

;$DICOMHEADER = $aArray[1]

if UBound(StringRegExp(String($DICOMHEADER), '(02000200)(.*?)', $STR_REGEXPARRAYMATCH))<1 then
   Return $sPatientName&"|"&string($sSex)&"|"&$sDOB&"|"&$sDoctorName ;exit because header wasn't found
EndIf

$offset=StringInStr($DICOMHEADER,"08009000")
$sDoctorName=BinaryToString("0x"&StringMid($DICOMHEADER,$offset+16,number("0x" & StringMid($DICOMHEADER,($offset+6*2),2))*2)) ;PHYSICIAN
$offset=StringInStr($DICOMHEADER,"10001000")
$sPatientName=BinaryToString("0x"&StringStripWS(StringMid($DICOMHEADER,$offset+16,number("0x" & StringMid($DICOMHEADER,($offset+6*2),2))*2),$STR_STRIPTRAILING)) ;PATIENT NAME
$offset=StringInStr($DICOMHEADER,"100020004C")
if $offset>0 then
$sID=(BinaryToString("0x"&StringMid($DICOMHEADER,$offset+16,number("0x" & StringMid($DICOMHEADER,($offset+6*2),2))*2))) ;PATIENT ID
;ConsoleWrite(" "&hex($offset))
;ConsoleWrite(" "&hex(StringMid($DICOMHEADER,($offset+6*2),2))*2)
Else
  $sID=""
EndIf

$offset=StringInStr($DICOMHEADER,"10003000")
$sDOB=(BinaryToString("0x"&StringMid($DICOMHEADER,$offset+16,number("0x" & StringMid($DICOMHEADER,($offset+6*2),2))*2))) ;PATIENT DOB
$offset=StringInStr($DICOMHEADER,"10004000")
$sSex=(BinaryToString("0x"&StringMid($DICOMHEADER,$offset+16,1*2))) ;PATIENT SEX
$offset=StringInStr($DICOMHEADER,"08005000")
$sAccession=(BinaryToString("0x"&StringMid($DICOMHEADER,$offset+16,number("0x" & StringMid($DICOMHEADER,($offset+6*2),2))*2))) ;ACCESSION

Return $sPatientName&"|"&string($sSex)&"|"&$sDOB&"|"&$sID&"|"&$sAccession&"|"&$sDoctorName
;Return $sPatientName&"|"&string($sSex)&"|"&$sDOB&"|"&$sDoctorName
EndFunc

Func _CopyDirWithProgress($sOriginalDir, $sDestDir)
  ;$sOriginalDir and $sDestDir are quite selfexplanatory...
  ;This func returns:
  ; -1 in case of critical error, bad original or destination dir
  ;  0 if everything went all right
  ; >0 is the number of file not copied and it makes a log file
  ;  if in the log appear as error message '0 file copied' it is a bug of some windows' copy command that does not redirect output...
Local $hTimer = TimerInit()

   If StringRight($sOriginalDir, 1) <> '\' Then $sOriginalDir = $sOriginalDir & '\'
   If StringRight($sDestDir, 1) <> '\' Then $sDestDir = $sDestDir & '\'
   If $sOriginalDir = $sDestDir Then Return -1

   ProgressOn('Copying...', 'Making list of files...' & @LF & @LF, '', -1, -1, 18)
   Local $aFileList = _FileSearch($sOriginalDir)
   If $aFileList[0] = 0 Then
      ProgressOff()
      SetError(1)
      Return -1
   EndIf

   If FileExists($sDestDir) Then
      If Not StringInStr(FileGetAttrib($sDestDir), 'd') Then
         ProgressOff()
         SetError(2)
         Return -1
      EndIf
   Else
      DirCreate($sDestDir)
      If Not FileExists($sDestDir) Then
         ProgressOff()
         SetError(2)
         Return -1
      EndIf
   EndIf

   Local $iDirSize, $iCopiedSize = 0, $fProgress = 0
   Local $c, $FileName, $iOutPut = 0, $sLost = '', $sError
   Local $Sl = StringLen($sOriginalDir)

   _Quick_Sort($aFileList, 1, $aFileList[0])

   $iDirSize = Int(DirGetSize($sOriginalDir) / 1024)

   ProgressSet(Int($fProgress * 100), $aFileList[$c], 'Coping file:')
   For $c = 1 To $aFileList[0]
      $FileName = StringTrimLeft($aFileList[$c], $Sl)
      ProgressSet(Int($fProgress * 100), $aFileList[$c] & ' -> '& $sDestDir & $FileName & @LF & 'Total KiB: ' & $iDirSize & @LF & 'Done KiB: ' & $iCopiedSize, 'Coping file:  ' & Round($fProgress * 100, 2) & ' %   ' & $c & '/' &$aFileList[0])

      If StringInStr(FileGetAttrib($aFileList[$c]), 'd') Then
         DirCreate($sDestDir & $FileName)
      Else
         If Not FileCopy($aFileList[$c], $sDestDir & $FileName, 1) Then
            If Not FileCopy($aFileList[$c], $sDestDir & $FileName, 1) Then;Tries a second time
               If RunWait(@ComSpec & ' /c copy /y "' & $aFileList[$c] & '" "' & $sDestDir & $FileName & '">' & @TempDir & '\o.tmp', '', @SW_HIDE)=1 Then;and a third time, but this time it takes the error message
                  $sError = FileReadLine(@TempDir & '\o.tmp',1)
                  $iOutPut = $iOutPut + 1
                  $sLost = $sLost & $aFileList[$c] & '  ' & $sError & @CRLF
               EndIf
               FileDelete(@TempDir & '\o.tmp')
            EndIf
         EndIf

         FileSetAttrib($sDestDir & $FileName, "+A-RSH");<- Comment this line if you do not want attribs reset.

         $iCopiedSize = $iCopiedSize + Int(FileGetSize($aFileList[$c]) / 1024)
         $fProgress = $iCopiedSize / $iDirSize
      EndIf
   Next

   ProgressOff()

   If $sLost <> '' Then;tries to write the log somewhere.
	    MsgBox($MB_SYSTEMMODAL, "Failed", "Copy Failed at "& Round(number(TimerDiff($hTimer)/1000),1)&" seconds")
      If FileWrite($sDestDir & 'notcopied.txt',$sLost) = 0 Then
         If FileWrite($sOriginalDir & 'notcopied.txt',$sLost) = 0 Then
            FileWrite(@WorkingDir & '\notcopied.txt',$sLost)
         EndIf
      EndIf
   Else
	    MsgBox($MB_SYSTEMMODAL, "Done", "Copy Completed in "& Round(number(TimerDiff($hTimer)/1000),1)&" seconds")
   EndIf


Return $iOutPut
EndFunc  ;==>_CopyDirWithProgress

Func _FileSearch($sIstr, $bSF = 1)
  ; $bSF = 1 means looking in subfolders
  ; $sSF = 0 means looking only in the current folder.
  ; An array is returned with the full path of all files found. The pos [0] keeps the number of elements.
;   Local $sIstr, $bSF, $sCriteria, $sBuffer, $iH, $iH2, $sCS, $sCF, $sCF2, $sCP, $sFP, $sOutPut = '', $aNull[1]
   Local  $sCriteria, $sBuffer, $iH, $iH2, $sCS, $sCF, $sCF2, $sCP, $sFP, $sOutPut = '', $aNull[1]
   $sCP = StringLeft($sIstr, StringInStr($sIstr, '\', 0, -1))
   If $sCP = '' Then $sCP = @WorkingDir & '\'
   $sCriteria = StringTrimLeft($sIstr, StringInStr($sIstr, '\', 0, -1))
   If $sCriteria = '' Then $sCriteria = '*.*'

  ;To begin we seek in the starting path.
   $sCS = FileFindFirstFile($sCP & $sCriteria)
   If $sCS <> - 1 Then
      Do
         $sCF = FileFindNextFile($sCS)
         If @error Then
            FileClose($sCS)
            ExitLoop
         EndIf
         If $sCF = '.' Or $sCF = '..' Then ContinueLoop
         $sOutPut = $sOutPut & $sCP & $sCF & @LF
      Until 0
   EndIf

  ;And after, if needed, in the rest of the folders.
   If $bSF = 1 Then
      $sBuffer = @CR & $sCP & '*' & @LF;The buffer is set for keeping the given path plus a *.
      Do
         $sCS = StringTrimLeft(StringLeft($sBuffer, StringInStr($sBuffer, @LF, 0, 1) - 1), 1);current search.
         $sCP = StringLeft($sCS, StringInStr($sCS, '\', 0, -1));Current search path.
         $iH = FileFindFirstFile($sCS)
         If $iH <> - 1 Then
            Do
               $sCF = FileFindNextFile($iH)
               If @error Then
                  FileClose($iH)
                  ExitLoop
               EndIf
               If $sCF = '.' Or $sCF = '..' Then ContinueLoop
               If StringInStr(FileGetAttrib($sCP & $sCF), 'd') Then
                  $sBuffer = @CR & $sCP & $sCF & '\*' & @LF & $sBuffer;Every folder found is added in the begin of buffer
                  $sFP = $sCP & $sCF & '\';                               for future searches
                  $iH2 = FileFindFirstFile($sFP & $sCriteria);         and checked with the criteria.
                  If $iH2 <> - 1 Then
                     Do
                        $sCF2 = FileFindNextFile($iH2)
                        If @error Then
                           FileClose($iH2)
                           ExitLoop
                        EndIf
                        If $sCF2 = '.' Or $sCF2 = '..' Then ContinueLoop
                        $sOutPut = $sOutPut & $sFP & $sCF2 & @LF;Found items are put in the Output.
                     Until 0
                  EndIf
               EndIf
            Until 0
         EndIf
         $sBuffer = StringReplace($sBuffer, @CR & $sCS & @LF, '')
      Until $sBuffer = ''
   EndIf

   If $sOutPut = '' Then
      $aNull[0] = 0
      Return $aNull
   Else
      Return StringSplit(StringTrimRight($sOutPut, 1), @LF)
   EndIf
EndFunc  ;==>_FileSearch

Func _Quick_Sort(ByRef $SortArray, $First, $Last);Larry's code
   Dim $Low, $High
   Dim $Temp, $List_Separator

   $Low = $First
   $High = $Last
   $List_Separator = StringLen($SortArray[ ($First + $Last) / 2])
   Do
      While (StringLen($SortArray[$Low]) < $List_Separator)
         $Low = $Low + 1
      WEnd
      While (StringLen($SortArray[$High]) > $List_Separator)
         $High = $High - 1
      WEnd
      If ($Low <= $High) Then
         $Temp = $SortArray[$Low]
         $SortArray[$Low] = $SortArray[$High]
         $SortArray[$High] = $Temp
         $Low = $Low + 1
         $High = $High - 1
      EndIf
   Until $Low > $High
   If ($First < $High) Then _Quick_Sort($SortArray, $First, $High)
   If ($Low < $Last) Then _Quick_Sort($SortArray, $Low, $Last)
EndFunc  ;==>_Quick_Sort

Func WM_NOTIFY($hWnd, $MsgID, $wParam, $lParam)
    Local $tagNMHDR, $event, $hwndFrom, $code
    $tagNMHDR = DllStructCreate("int;int;int", $lParam)
    If @error Then Return 0
    $code = DllStructGetData($tagNMHDR, 3)
    ;If $wParam = $lv And $code = $NM_DBLCLK Then $DoubleClicked = True
    ;If $wParam = $lv And $code = $NM_CLICK Then $DoubleClicked = True
	  Switch $wParam
		 case $lv
             Switch DllStructGetData($tagNMHDR, 3)
				case $NM_DBLCLK
                 $DoubleClicked = True
				case $NM_CLICK
                 $sSelectedFile=StringRegExp(_GUICtrlListView_GetItemTextString($lv), '([^|]+$)', $STR_REGEXPARRAYMATCH)
				 If UBound($sSelectedFile)>0 Then
				 $sSelectedFile=$sSelectedFile[0]
				 EndIf
			  EndSwitch
		  If StringRight(GUICtrlRead($input1),1) = @CR Then ;;DEAD CODE DOESNT WORK
                Send("{BS}")
                ;ConsoleWrite("J")
            EndIf

		 EndSwitch

    Return $GUI_RUNDEFMSG
EndFunc

Func _Exit()
	;_GUICtrlListView_UnRegisterSortCallBack($hListView2)
	;_SaveOptions()
	GUIDelete();
	Exit 0
EndFunc