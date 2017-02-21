'**************************************************
' Script Name: ImportToExcel.vbs -- Ver3.0
' Description: This script will extract function names and their parameters from a vbs library file, 
' 	       and store it in excel.
'**************************************************

'**************************************************
'Generic MS Excel VBA constants
'**************************************************

'Excel cell insertion shift direction constants
Const xlShiftDown = -4121
Const xlShiftToRight = -4161

'Excel cell deletion shift direction constants
Const xlShiftToLeft = -4159
Const xlShiftUp = -4162

'Excel horizontal alignment constants
Const xlHAlignCenter = -4108
Const xlHAlignCenterAcrossSelection = 7
Const xlHAlignDistributed = -4117
Const xlHAlignFill = 5
Const xlHAlignGeneral = 1
Const xlHAlignJustify = -4130
Const xlHAlignLeft = -4131
Const xlHAlignRight = -4152

'Excel vertical alignment constants
Const xlVAlignBottom = -4107
Const xlVAlignCenter = -4108
Const xlVAlignDistributed = -4117
Const xlVAlignJustify = -4130
Const xlVAlignTop = -4160

'Excel summary row constants
Const xlSummaryOnLeft = -4131
Const xlSummaryOnRight = -4152
Const xlSummaryAbove = 0
Const xlSummaryBelow = 1

'Excel border index constants
Const xlDiagonalDown = 5
Const xlDiagonalUp = 6
Const xlEdgeBottom = 9
Const xlEdgeLeft = 7
Const xlEdgeRight = 10
Const xlEdgeTop = 8
Const xlInsideHorizontal = 12
Const xlInsideVertical = 11

'Excel border weight constants
Const xlHairline = 1
Const xlMedium = -4138
Const xlThick = 4
Const xlThin = 2

'Excel line style constants
Const xlContinuous = 1
Const xlDash = -4115
Const xlDashDot = 4
Const xlDashDotDot = 5
Const xlDot = -4118
Const xlDouble = -4119
Const xlLineStyleNone = -4142
Const xlSlantDashDot = 13 

'**************************************************
'Declare functions
'**************************************************

Public Function fGetVBSFile(strFileType)
'Set objDialog = CreateObject("UserAccounts.CommonDialog")
	
	Select Case UCase(strFileType)
	Case "TEXT"
		objDialog.Filter = "Text Files|*.txt"
	Case "EXCEL"
		fGetVBSFile = "C:\Gentera\Book1.xls"
		'objDialog.Filter = "Excel Files|*.xls"
	Case "VBS"
		'fGetVBSFile = "C:\Gentera\FunctionLibrary_Credito - Copy.txt"
		fGetVBSFile = "C:\Gentera\FunctionLibrary_Credito.vbs"
		'objDialog.Filter = "VBS Files|*.vbs"
	End Select
	
	'objDialog.FilterIndex = 1
	'objDialog.InitialDir = "C:\"
			
	'intResult = objDialog.ShowOpen
	'If intResult = True Then
	'	fGetVBSFile = CStr(objDialog.FileName)
	'Else
	'	fGetVBSFile = False
	'End If
	
	'Set objDialog = Nothing
End Function

Public Function RegExpTest(strInput,strMatchPattern)
	Dim objRegEx, Match, Matches
	Set objRegEx = New RegExp 
 
    	objRegEx.Global = True
    	objRegEx.Pattern = strMatchPattern 
 
    	Set Matches = objRegEx.Execute(strInput) 
		If Matches.Count > 0 Then
			RegExpTest = Matches(0).FirstIndex+1
		Else
			RegExpText = 0
		End If
End Function

'**************************************************
'Main script
'**************************************************

'Open the VBS file
Dim objFso, objFile
Dim objXLApp,objXLWkBooks,objXLWkBook,objXLWkSheet

'Set file object for the vbs files
Set objFso = CreateObject("Scripting.FileSystemObject")

'Set file object for the excel file
Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = False
Set objXLWkBooks = objXLApp.Workbooks

'Prompt user for Input VBS file
rc = Msgbox("Please Select the VBS function library file.",vbOKCancel,"Import Functions to Excel")
If rc = 2 Then
	Msgbox "Exiting Script Execution!",vbOKOnly + vbCritical,"Import Functions to Excel"
	Set objFile = Nothing
	Set objFSO = Nothing
	WScript.Quit(1)
End If 

'Get the file name from the File open dialog box
strFilePath = fGetVBSFile("VBS")

'Prompt user for Excel Reference file
rc = Msgbox("Please Select the Excel file which will import the VBS function library.",vbOKCancel,"Import Functions to Excel")
If rc = 2 Then
	Msgbox "Exiting Script Execution!",vbOKOnly + vbCritical,"Import Functions to Excel"
	Set objFile = Nothing
	Set objFSO = Nothing
	WScript.Quit(1)
End If

'Get the file name from the File open dialog box
strExcelFilePath = fGetVBSFile("EXCEL")


If strFilePath <> False And strExcelFilePath <> False Then
	strFileName = Mid(Mid(strFilePath,1,InStrRev(UCase(strFilePath),".VBS")-1),InStrRev(Mid(strFilePath,1,InStrRev(UCase(strFilePath),".VBS")-1),"\")+1)
	strExcelFileName = Mid(Mid(strExcelFilePath,1,InStrRev(UCase(strExcelFilePath),".XLS")+3),InStrRev(Mid(strExcelFilePath,1,InStrRev(UCase(strExcelFilePath),".XLS")+3),"\")+1)
	strFileNameOrig = strFileName
	strSheetNameOrig = strFileNameOrig
	
	'Open the VBS file
	Set objFile = objFso.OpenTextFile(strFilePath,1,False,True)
	
	'Open the Excel file
	Set objXLWkBook = objXLWkBooks.Open(strExcelFilePath)
	
	Set objXLWkSheet = objXLWkBook.Worksheets("Function Tracker")
	objXLWkSheet.Select
	
	If Len(strFileName) > 31 Then
		'Check if the file Entry is present in the Function Import History
		blnEntryExistsInFIH = False
		blnSheetExists = False
		
		'Find Last data row
		iLastRow = 5
		For iCnt = 6 To 262
			If Trim(objXLWkSheet.Cells(iCnt,3).Value) = "" And Trim(objXLWkSheet.Cells(iCnt+1,3).Value) = "" Then
				iLastRow = iCnt-1
				Exit For
			End If
		Next
		
		iFileRow = 0
		For iCnt = 6 To iLastRow
			If UCase(objXLWkSheet.Cells(iCnt,3).Value) = UCase(strFileNameOrig&".vbs") Then
				blnEntryExistsInFIH = True
				iFileRow = iCnt
				Exit For
			End If
		Next
		
		'Check if a sheet to be added already exists
		blnSheetExistsTmp = False
		If blnEntryExistsInFIH = True Then
			For iCnt = 2 To objXLWkBook.Worksheets.Count
				If UCase(objXLWkBook.Worksheets(iCnt).Name) = UCase(objXLWkSheet.Cells((iFileRow),2).Value) Then	'Start with 2nd Sheet and 6th row of Function_Tracker
					blnSheetExistsTmp = True
					Exit For
				End If
			Next
		End If
		
		'If the sheet does not exist, then name it as <sheet name>$<num>, where <num> is the one greater than the last number with similar name.
		'For example, if the vbs file name is ABC_DEF_GHI_JKL_MNO_PQR.vbs, and there are already sheets named ABC_DEF_GHI_JKL_MNO_P$1, and ABC_DEF_GHI_JKL_MNO_P$2,
		'the the required sheet name should be ABC_DEF_GHI_JKL_MNO_P$3
		If blnSheetExistsTmp = False Then
			iMax = 0
			iMaxTmp = 0
			For iCnt = 6 To iLastRow
				If UCase(Mid(objXLWkSheet.Cells(iCnt,2).Value,1,29)) = UCase(Mid(strFileNameOrig,1,29)) Then
					iMaxTmp = CInt(Mid(objXLWkSheet.Cells(iCnt,2).Value,31))
					If iMaxTmp > iMax Then
						iMax = iMaxTmp
					End If
					Exit For
				End If
			Next
			strFileName = Mid(strFileName,1,29)&"$"&(iMax+1)
			blnSheetExists = False
		Else
			'If the sheet exists, then name it as the existing sheet name
			strFileName = objXLWkSheet.Cells(iFileRow,2).Value
			blnSheetExists = True
		End If
	Else
		'Check if the worksheet is already added to the given workbook
		blnSheetExists = False
		For iCnt = 1 To objXLWkBook.Worksheets.Count
			If objXLWkBook.Worksheets(iCnt).Name = strFileName Then
				blnSheetExists = True
				Exit For
			End If
		Next
	End If
	
	If blnSheetExists = True Then
		'Check if user wants to overwrite the sheet or create another one.
		rc = Msgbox("The given Excel file already contains the sheet "& strFileName & "." & vbCRLF &_
			    "Do you want to overwrite the sheet with fresh import?",vbYesNo + vbExclamation,"Import Functions to Excel")
		
		If rc = 6 Then
			objXLApp.DisplayAlerts = False
			objXLWkBook.Worksheets(strFileName).Delete
			objXLApp.DisplayAlerts = True
			objXLWkBook.Save
			
			blnSheetExists = False
		ElseIf rc = 7 Then
			objXLWkBook.Close

			objXLWkBooks.Close
			objXLApp.Quit

			Msgbox "Import Aborted!",vbOKOnly + vbCritical,"Import Functions to Excel"
			
			Set objXLWkSheet = Nothing
		End If
	End If
	
	If blnSheetExists = False Then
		objXLWkBook.Worksheets.Add.Name = strFileName
		Set objXLWkSheet = objXLWkBook.Worksheets(strFileName)
		objXLWkSheet.Select
		strSheetName = objXLWkSheet.Name
		
		'Move Sheet to End
		objXLWkSheet.Move ,objXLWkBook.Worksheets(objXLWkBook.Worksheets.Count)
		
		'Add the column headers
		objXLWkSheet.Cells(1,1).Value = "Function Name"
		objXLWkSheet.Cells(1,2).Value = "Function Description"
		objXLWkSheet.Cells(1,3).Value = "Input Parameters"
		objXLWkSheet.Cells(1,4).Value = "Parameter Description"
		objXLWkSheet.Cells(1,5).Value = "Output Parameters"
		
		'Common properties for all cells
		objXLWkSheet.Range("A1:IV65536").Select
		With objXLWkSheet.Range("A1:IV65536")
			.Font.Name = "Arial"
			.Font.Size = 8
			.Interior.ColorIndex = 2
		End With
		
		'Format the column header
		objXLWkSheet.Columns("A").Select
		With objXLWkSheet.Columns("A")
			.ColumnWidth = 30
			.Font.Bold = True
		End With
		
		objXLWkSheet.Columns("B").Select
		With objXLWkSheet.Columns("B")
			.ColumnWidth = 45
			.Font.Bold = False
		End With
		
		objXLWkSheet.Columns("C").Select
		With objXLWkSheet.Columns("C")
			.ColumnWidth = 20
			.Font.Bold = True
		End With
		
		objXLWkSheet.Columns("D").Select
		With objXLWkSheet.Columns("D")
			.ColumnWidth = 40
			.Font.Bold = False
		End With
		
		objXLWkSheet.Columns("E").Select
		With objXLWkSheet.Columns("E")
			.ColumnWidth = 30
			.Font.Bold = False
		End With
		
		objXLWkSheet.Range("A1:E1").Select
		With objXLWkSheet.Range("A1:E1")
			.WrapText = True
			.Font.Name = "Arial"
			.Font.Size = 10
			.Font.Bold = True
			.Interior.ColorIndex = 36	'LightYellow
			.Borders.LineStyle = xlContinuous
			.Borders.ColorIndex = 1		'Black
		End With
		
		objXLWkSheet.Range("A1").Select
	End If
	
	If blnSheetExists = False Then
		On Error Resume Next
	
		Dim iCurrRow,iStartRow,iEndRow
		Dim arrLine(30000),iArrIndex
		
		iArrIndex = 0
		iCurrRow = 2
		'ReadLine the input file to locate the functions/subs
		Do Until objFile.AtEndOfStream
			strLine = objFile.Readline
			
			If Trim(strLine) <> "" Then
				arrLine(iArrIndex) = strLine
					
				'Increment array index
				iArrIndex = iArrIndex + 1
			End If
			
			If Instr(1,UCase(strLine),"FUNCTION") > 0 Or Instr(1,UCase(strLine),"PUBLIC SUB") > 0 Then
			
				'Extract the function description and parameter description
				Dim blnKeepLooking,iTmpIndex,iIndex,blnInnerLoop,blnFuncDescBorder,blnFoundAgain
				Dim strFName,strDesc,strInParams,strOutParams
				
				blnKeepLooking = True
				iTmpIndex = iArrIndex - 2
				
				strFName     = ""
				strDesc      = ""
				strInParams  = ""
				strOutParams = ""
				
				blnFuncDescBorder = False
				blnFoundAgain = False

				'If the current iTmpIndex points to the '#*****...***** line then move pointer up by 1 line
				'If Instr(1,Trim(UCase(arrLine(iTmpIndex))),"**********") > 0 Then
				'	iTmpIndex = iTmpIndex - 1
				'End If
				
				
				While blnKeepLooking = True
					'Look for the '#*****...***** line. First time it appears, that means its the lower border of function description
					'Second time it should be the upper border of the function description
					If Instr(1,Trim(UCase(arrLine(iTmpIndex))),"**********") > 0 And blnFuncDescBorder = False Then
						blnFuncDescBorder = True
						iTmpIndex = iTmpIndex - 1
					End If
					
					If Instr(1,Trim(UCase(arrLine(iTmpIndex))),"RETURN VALUES") > 0 Then
						strOutParams = Trim(Mid(arrLine(iTmpIndex),Instr(1,arrLine(iTmpIndex),":")+1))
						iIndex = iTmpIndex+1
						blnInnerLoop = True
						While blnInnerLoop = True
							If Instr(1,Trim(UCase(arrLine(iIndex))),"**********") > 0 Then
								blnInnerLoop = False
							Else
								strOutParams = strOutParams & vbLF & Trim(Mid(arrLine(iIndex),RegExpTest(arrLine(iIndex),"[A-Za-z0-9]"))) 
								iIndex = iIndex + 1
							End If	
						WEnd
					ElseIf Instr(1,Trim(UCase(arrLine(iTmpIndex))),"INPUT PARAMETERS") > 0 Then
						strInParams = Trim(Mid(arrLine(iTmpIndex),Instr(1,arrLine(iTmpIndex),":")+1))
						iIndex = iTmpIndex+1
						blnInnerLoop = True
						While blnInnerLoop = True
							If Instr(1,Trim(UCase(arrLine(iIndex))),"RETURN VALUES") > 0 Then
								blnInnerLoop = False
							Else
								strInParams = strInParams & vbLF & Trim(Mid(arrLine(iIndex),RegExpTest(arrLine(iIndex),"[A-Za-z0-9]"))) 
								iIndex = iIndex + 1
							End If	
						WEnd
					ElseIf Instr(1,Trim(UCase(arrLine(iTmpIndex))),"DESCRIPTION") > 0 Then
						strDesc = Trim(Mid(arrLine(iTmpIndex),Instr(1,arrLine(iTmpIndex),":")+1))
						iIndex = iTmpIndex+1
						blnInnerLoop = True
						While blnInnerLoop = True
							If Instr(1,Trim(UCase(arrLine(iIndex))),"INPUT PARAMETERS") > 0 Then
								blnInnerLoop = False
							Else
								strDesc = strDesc & vbLF & Trim(Mid(arrLine(iIndex),RegExpTest(arrLine(iIndex),"[A-Za-z0-9]"))) 
								iIndex = iIndex + 1
							End If	
						WEnd
					ElseIf Instr(1,Trim(UCase(arrLine(iTmpIndex))),"FUNCTION NAME") > 0 Then
						strFName = Trim(Mid(arrLine(iTmpIndex),Instr(1,arrLine(iTmpIndex),":")+1))
						iIndex = iTmpIndex+1
						blnInnerLoop = True
						While blnInnerLoop = True
							If Instr(1,Trim(UCase(arrLine(iIndex))),"DESCRIPTION") > 0 Then
								blnInnerLoop = False
							Else
								strFName = strFName & vbLF & Trim(Mid(arrLine(iIndex),RegExpTest(arrLine(iIndex),"[A-Za-z0-9]"))) 
								iIndex = iIndex + 1
							End If	
						WEnd
					End If 
					
					If Instr(1,Trim(UCase(arrLine(iTmpIndex))),"END FUNCTION") > 0 Or Instr(1,Trim(UCase(arrLine(iTmpIndex))),"END SUB") > 0 Or iTmpIndex = 0 Then
						blnKeepLooking = False
					ElseIf Instr(1,Trim(UCase(arrLine(iTmpIndex))),"**********") > 0 Then
						blnKeepLooking = False
					Else
						iTmpIndex = iTmpIndex - 1
					End If
				WEnd
				
				'Extract the Function Name
				strFuncName = Trim(Mid(Mid(strLine,1,Instr(1,strLine,"(")-1),InStrRev(Trim(Mid(strLine,1,Instr(1,strLine,"(")-1))," ")+1))
				
				objXLWkSheet.Cells(iCurrRow,1).Value = Trim(strFuncName)
				With objXLWkSheet.Cells(iCurrRow,1)
					.Interior.ColorIndex = 15	'LightGray
					.VerticalAlignment = xlVAlignTop
					.Borders.LineStyle = xlContinuous
					.Borders.ColorIndex = 1		'Black
				End With
				
				'Extract the Function Parameters
				strFuncParams = Trim(Mid(strLine,Instr(1,strLine,"(")+1,(Instr(1,strLine,")"))-(Instr(1,strLine,"(")+1)))
				iStartRow = iCurrRow
				If strFuncParams <> "" Then
					arrFuncParams = Split(strFuncParams,",")
												
					For iCnt = 0 To UBound(arrFuncParams)
						objXLWkSheet.Cells(iCurrRow,3).Value = Trim(arrFuncParams(iCnt))
						With objXLWkSheet.Cells(iCurrRow,3)
							.Font.ColorIndex = 5		'Blue
							.VerticalAlignment = xlVAlignTop
							.Borders.LineStyle = xlContinuous
							.Borders.ColorIndex = 1		'Black
						End With
						iCurrRow = iCurrRow + 1
					Next
				Else
					objXLWkSheet.Cells(iCurrRow,3).Value = "< None >"
					With objXLWkSheet.Cells(iCurrRow,3)
						.Font.ColorIndex = 5		'Blue
						.VerticalAlignment = xlVAlignTop
						.Borders.LineStyle = xlContinuous
						.Borders.ColorIndex = 1		'Black
					End With
					iCurrRow = iCurrRow + 1
				End If
				
				iEndRow = iCurrRow - 1
				If iEndRow < iStartRow Then
					iEndRow = iStartRow
				End If
				
				'Set value for Function Description and Merge the Cells for Function Description
				objXLWkSheet.Cells(iStartRow,2).Value = "["& strFName &"] "& strDesc
				objXLWkSheet.Range("B"&iStartRow&":B"&iEndRow).Select
				With objXLWkSheet.Range("B"&iStartRow&":B"&iEndRow)
					.MergeCells = True
					.WrapText = True
					.VerticalAlignment = xlVAlignTop
					.Borders.LineStyle = xlContinuous
					.Borders.ColorIndex = 1		'Black
				End With
				
				'Set value for Parameter Description and Merge the Cells for Parameter Description
				objXLWkSheet.Cells(iStartRow,4).Value = strInParams
				objXLWkSheet.Range("D"&iStartRow&":D"&iEndRow).Select
				With objXLWkSheet.Range("D"&iStartRow&":D"&iEndRow)
					.MergeCells = True
					.WrapText = True
					.VerticalAlignment = xlVAlignTop
					.Borders.LineStyle = xlContinuous
					.Borders.ColorIndex = 1		'Black
				End With
				
				'Set value for Output Parameters and Merge the Cells for Output Parameters
				objXLWkSheet.Cells(iStartRow,5).Value = strOutParams
				objXLWkSheet.Range("E"&iStartRow&":E"&iEndRow).Select
				With objXLWkSheet.Range("E"&iStartRow&":E"&iEndRow)
					.MergeCells = True
					.WrapText = True
					.VerticalAlignment = xlVAlignTop
					.Borders.LineStyle = xlContinuous
					.Borders.ColorIndex = 1		'Black
				End With
				
				'Create a blank row
				iCurrRow = iCurrRow + 1
			End If
		Loop
		
		'Freeze Panes on column B and row 1
		objXLWkSheet.Range("B2").Select
		objXLApp.ActiveWindow.FreezePanes = True

		'Set the selection back to A1
		objXLWkSheet.Range("A1").Select
		
		'Update Function Tracker worksheet
		Set objXLWkSheet = objXLWkBook.Worksheets("Function Tracker")
		objXLWkSheet.Select
		
		'Check if the file name entry is already present, update that row
		iFuncRow = -1
		iLastRow = 5
		For iCnt = 6 To 262
			If Trim(objXLWkSheet.Cells(iCnt,3).Value) = "" And Trim(objXLWkSheet.Cells(iCnt+1,3).Value) = "" Then
				iLastRow = iCnt-1
				Exit For
			End If
		Next
		
		For iCnt = 6 To iLastRow
			If UCase(objXLWkSheet.Cells(iCnt,3).Value) = UCase(strFileNameOrig&".vbs") Then
				iFuncRow = iCnt
				Exit For
			End If
		Next
		
		If iFuncRow = -1 Then
			iFuncRow = iLastRow+1
		End If
		
		'File HyperLink Name (B)
		objXLWkSheet.Cells(iFuncRow,2).Value = strSheetName
		
		'File Name (C)
		objXLWkSheet.Cells(iFuncRow,3).Value = strFileNameOrig & ".vbs"
		
		'Last Update Date (D)
		objXLWkSheet.Cells(iFuncRow,4).Value = Now()
		
		'Create Link to the sheet
		objXLWkSheet.Hyperlinks.Add objXLWkSheet.Range("E"&iFuncRow),strExcelFileName,"'"&strSheetName&"'!A1",strSheetNameOrig,"Click Here"
		
		objXLWkSheet.Range("C"&iFuncRow&":E"&iFuncRow).Select
		With objXLWkSheet.Range("C"&iFuncRow&":E"&iFuncRow)
			.Font.Name = "Arial"
			.Font.Size = 8
			.Font.Bold = True
			.Interior.ColorIndex = 19	'
			.WrapText = True
			.VerticalAlignment = xlVAlignTop
			.Borders.LineStyle = xlContinuous
			.Borders.ColorIndex = 1		'Black
		End With
		objXLWkSheet.Range("C2:E2").Select
	End If
	
	If blnSheetExists = False Then
		objXLWkBook.Save
		objXLWkBook.Close

		objXLWkBooks.Close
		objXLApp.Quit
		
		Msgbox "Import Complete!",vbOKOnly,"Import Functions to Excel"

		Set objXLWkSheet = Nothing
	End If
	
	Set objFile = Nothing
	Set objFSO = Nothing
	
	Set objXLWkBook = Nothing
	Set objXLWkBooks = Nothing
	Set objXLApp = Nothing
End If

