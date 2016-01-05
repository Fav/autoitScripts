#Region ;**** ���������� ACNWrapper_GUI ****
#AutoIt3Wrapper_Outfile=sql������.exe
#AutoIt3Wrapper_Outfile_x64=sql����-x64.exe
#AutoIt3Wrapper_Compile_Both=y
#EndRegion ;**** ���������� ACNWrapper_GUI ****
#include <Excel.au3>
#include<array.au3>
#Include <GuiComboBox.au3>

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=c:\autoit\src\common\form1.kxf
$Form1_1 = GUICreate("Oracle����ű�����", 514, 269, 192, 124)
$Label1 = GUICtrlCreateLabel("excel·��", 72, 32, 60, 17)
$Input1 = GUICtrlCreateInput("", 136, 32, 233, 21)
$Button1 = GUICtrlCreateButton("...", 384, 32, 33, 25)
$Combo1 = GUICtrlCreateCombo("", 136, 72, 233, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$Button2 = GUICtrlCreateButton("��ʼ", 360, 168, 65, 65)
$Label2 = GUICtrlCreateLabel("���Ŀ¼", 72, 104, 52, 17)
$Input2 = GUICtrlCreateInput( @DesktopDir, 136, 104, 233, 21)
$Button3 = GUICtrlCreateButton("...", 384, 104, 33, 25)
$Label3 = GUICtrlCreateLabel("��    ��", 72, 72, 60, 17)
$Label4 = GUICtrlCreateLabel("", 80, 192, 36, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $Button1
			$file = GetFilePath()
			SetInputText($Input1,$file)
			Local $oExcel = _ExcelBookOpen($file, 0)
			$ls = _ExcelSheetList($oExcel)
			If UBound($ls)>0 Then
				_GUICtrlComboBox_ResetContent($Combo1)
				For $i = 1 To UBound($ls)-1
					_GUICtrlComboBox_AddString($Combo1,$ls[$i])
				Next
			EndIf
			_GUICtrlComboBox_SetCurSel($Combo1, 0)
			_ExcelBookClose($oExcel)
		Case $Button3
			SetInputText($Input2,GetDir())
		Case $Button2
			OutPut(GUICtrlRead($Input1),GUICtrlRead($Combo1),GUICtrlRead($Input2))
	EndSwitch
WEnd

Func SetInputText($Input,$text)
	If $text <> "" Then
		GUICtrlSetData($Input,$text,"")
	EndIf
EndFunc

Func GetFilePath()
	Local $var = FileOpenDialog("", @WindowsDir & "\", "Excel (*.xls;*.xlsx)", 1)
	If @error Then
		MsgBox(4096,"","û��ѡ���ļ�!")
	EndIf
	Return $var
EndFunc

Func GetDir()
	Local $var =FileSelectFolder("ѡ��Ŀ¼.", @DesktopDir)
	If @error Then
		MsgBox(4096,"","û��ѡ��Ŀ¼!")
	EndIf
	Return $var
EndFunc


Func OutPut($filePath,$sheetName,$outdir)
	;Local $oExcel = _ExcelBookNew() ;Create new book, make it visible
	;Local $filePath = "C:\project\henan\2015-12-26\����ˮˮ��.xlsx"
	;Local $sheetName = '����ˮ���Կ��ṹ'
	Local $outFileName = $sheetName & ".txt"
	Local $oExcel = _ExcelBookOpen($filePath, 0)
	Local $outFilePath = $outdir & "\" &$outFileName
	Local $oFile = FileOpen($outFilePath,9)
	
	If @error = 1 Then
		MsgBox(0, "����!", "�޷���������!")
		Exit
	ElseIf @error = 2 Then
		MsgBox(0, "����!", "�ļ�������!")
		Exit
	EndIf

	;$iNumberOfWorksheets = $oExcel.Worksheets.Count
	;MsgBox(0, "", $oExcel.Worksheets.Count)
	
	_ExcelSheetActivate($oExcel, $sheetName)
	;ѭ����ȡ
	$ix = 1
	
	$content = _ExcelReadCell($oExcel, $ix, 1)
	While $content <> ""
		$dTableDescript = _ExcelReadCell($oExcel, $ix, 1) ;��ɽ������Ϣ��
		FileWrite($oFile,"--" & $dTableDescript & @CRLF)
		$dTableName = _ExcelReadCell($oExcel, $ix + 1, 1) ;KDHS01A
		
		;д���ı�
		FileWrite($oFile,"CREATE TABLE " & $dTableName & "(" & @CRLF)
		
		ConsoleWrite($dTableName & @CRLF)	
		Dim $primeKeyArr[1]=['1']
		
		Dim $fieldCreateArr[1]=['1']
		
		Dim $fieldDesArr[1]=['comment on table '& $dTableName &" is '" & $dTableDescript & "'"]
		
		;д������Ϣ & "
		$fieldCount = 0
		$fieldInfo = _ExcelReadCell($oExcel, $ix+3+$fieldCount, 1)
		While $fieldInfo <> ""
			$fieldNameCN = _ExcelReadCell($oExcel, $ix+3+$fieldCount, 2)
			$fieldName = _ExcelReadCell($oExcel, $ix+3+$fieldCount, 3)
			$fieldType = _ExcelReadCell($oExcel, $ix+3+$fieldCount, 4)
			ConsoleWrite($fieldInfo & @CRLF)
			GUICtrlSetData($Label4,$fieldInfo,"")
			;����
			If _ExcelReadCell($oExcel, $ix+3+$fieldCount, 5) == '��' Then
				_ArrayAdd($primeKeyArr, $fieldName)
			EndIf
			$isNullDes = ''
			If _ExcelReadCell($oExcel, $ix+3+$fieldCount, 6) == '��' Then
				$isNullDes = " not null"
			EndIf
			$fieldDes = _ExcelReadCell($oExcel, $ix+3+$fieldCount, 7)
			If $fieldDes <> "" Then
				$fieldDes = "("& $fieldDes & ")"
			EndIf
			
			_ArrayAdd($fieldCreateArr,"    " & $fieldName&" "& $fieldType & $isNullDes);�ĸ��ո�
			;ConsoleWrite("    " & $fieldName&" "& $fieldType & $isNullDes &@CRLF)
			_ArrayAdd($fieldDesArr,"comment on column " & $dTableName & "." & $fieldName & " is '" &$fieldNameCN & $fieldDes &"'")
			
			$fieldCount +=1
			$fieldInfo = _ExcelReadCell($oExcel, $ix+3+$fieldCount, 1)
		WEnd
		_ArrayDelete($primeKeyArr,0)
		_ArrayDelete($fieldCreateArr,0)
		;д���ı�-�ֶ�
		FileWrite($oFile,_ArrayToString($fieldCreateArr,"," & @CRLF) )
		
		$primeKeyDes = _ArrayToString($primeKeyArr,"��")
		If $primeKeyDes <> "" Then
			$primeKeyDes ="," & @CRLF &"    primary key (" & $primeKeyDes & ")"
		EndIf

		;д���ı�-����
		FileWrite($oFile,$primeKeyDes) 
		;д���ı�
		FileWrite($oFile,@CRLF & ");" & @CRLF) 
		;д���ı�-�ֶ�
		FileWrite($oFile,_ArrayToString($fieldDesArr,";" & @CRLF))
		

		$ix +=$fieldCount+5
		$content = _ExcelReadCell($oExcel, $ix, 1)
		;д���ı�
		FileWrite($oFile,";" & @CRLF& @CRLF) 
	WEnd

	;MsgBox(0, "Exiting", "Notice How Sheet2 is Active and not Sheet1" & @CRLF & @CRLF & "Now Press OK to Save File and Exit")
	;_ExcelBookSaveAs($oExcel, @TempDir & "\Temp.xls", "xls", 0, 1) ; Now we save it into the temp directory; overwrite existing file if necessary
	_ExcelBookClose($oExcel) ; And finally we close out
	FileClose($oFile)
	MsgBox(0,"���","���")
EndFunc   ;==>_Main
