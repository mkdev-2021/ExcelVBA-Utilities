'-----------------------------------------------------------------------------------------------------------------------------------------------------
Sub CopySheetsToNewWorkBook(ByRef WS_Names As Variant, new_WB_Name As String)

    Dim wbOld As Workbook
    Dim wbNew As Workbook
    Dim firstSheetToCopy As Boolean
    
    Dim tmp_WSName As String
    
    Set wbOld = ActiveWorkbook
    
    firstSheetToCopy = True
    For i = 0 To UBound(WS_Names)
    
        tmp_WSName = WS_Names(i)
        If WorksheetExists(tmp_WSName) Then
        
            If firstSheetToCopy Then
                wbOld.Sheets(tmp_WSName).Copy
                Set wbNew = ActiveWorkbook
                firstSheetToCopy = False
            Else
                wbOld.Sheets(tmp_WSName).Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
            End If
            
        End If
        
    Next i

    Dim fldrName As String
    fldrName = wbOld.Path
    
    wbNew.SaveAs Filename:=fldrName & "\" & new_WB_Name & ".xlsx", _
        FileFormat:=51, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    
    wbNew.Close

    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
    
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ArrayToRange()
	Dim arrayStr As Variant
    Dim Rng As Range
  
    Set Rng = WD5_Rpt_WS.Range("WBS_Unique_List").Resize(UBound(arrayStr))
    Rng = Application.Transpose(arrayStr)
    
   '------- sort the range ---------
    Rng.Sort key1:=Rng, order1:=xlAscending, MatchCase:=False

End sub
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub SortColumn()
    '---5. Sort the rows based on a column C with expanded range---
    ipRowCnt = temp_WS.Range("A" & Rows.Count).End(xlUp).Row
    
    Application.StatusBar = "Sort rows based on Data & time...."
    temp_WS.Sort.SortFields.Clear
    temp_WS.Sort.SortFields.Add Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    With temp_WS.Sort
        .SetRange Range("A2:C" & ipRowCnt)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'' Get the Column number based on the input columnname

Function GetColumn(vWorksheet As Worksheet, ColumnName As String, Optional ByVal ColumnNamesStartCell As Variant) As Integer

    Dim tmpWS As Worksheet
    Dim ColCounter As Long, i As Long, StartCell As String
    
    On Error Resume Next
    Set tmpWS = vWorksheet
    If tmpWS Is Nothing Then
        GetColumn = -1
    Else
        tmpWS.Activate
        
        If IsMissing(ColumnNamesStartCell) Then
            StartCell = "A1"
        Else
            StartCell = ColumnNamesStartCell
        End If
        
        With Range(StartCell)
            ColCounter = Range(.Offset(0, 0), .End(xlToRight)).Columns.Count
            For i = 0 To ColCounter - 1
                If LCase(.Offset(0, i)) = LCase(ColumnName) Then
                    GetColumn = i
                    Exit Function
                End If
            Next i
        End With
        
        'Column name not found
        GetColumn = -1
    End If
    
End Function


''Function to Re-arrange columns in the order specified in an array
Function ReArrangeColumns(vWorksheet As Worksheet, ByRef Col_Names_Array As Variant, Optional Delete_Other_Rows As Boolean) As Integer
    
    Dim Col_Num As Integer
    Dim Col_Name As String
    Dim Count_Dwn As Integer
    Dim MaxCol As Integer
    
    On Error GoTo ErrorHandler
        
    Count_Dwn = UBound(Col_Names_Array)
    
    vWorksheet.Activate
    For i = 0 To UBound(Col_Names_Array)
        
        Col_Name = Col_Names_Array(Count_Dwn)
        Col_Num = GetColumn(vWorksheet, Col_Name) + 1
        
        If Col_Num > 0 Then         ' Column heading found
          If Col_Num > 1 then
            vWorksheet.Columns(Col_Num).Select
            Selection.Cut
            Columns("A:A").Select
            Selection.Insert shift:=xlToRight
	   End if
        Else                        ' Column heading not found -> insert empty column
            vWorksheet.Columns(1).Select
            Selection.Insert shift:=xlToRight
            Range("A1").Value = "No Data - " & Col_Name
        End If
        Count_Dwn = Count_Dwn - 1
    Next i
    
    'Now Delete all the other extra columns in the sheet
    If Delete_Other_Rows Then
        vWorksheet.Activate
        MaxCol = vWorksheet.UsedRange.Columns.Count
'        vWorksheet.Range(Cells(1, UBound(Col_Names_Array) + 1), Cells(1, MaxCol)).Select
        vWorksheet.Range(Cells(1, UBound(Col_Names_Array) + 1), Cells(1, MaxCol)).EntireColumn.Delete
    End If
    
    ReArrangeColumns = 0
    Exit Function

ErrorHandler:
    ReArrangeColumns = 1

End Function


-----------------------------------------------------------------------------------------------------------------------------------------------------
''Get the Sheet name
Function SheetName(rCell As Range, Optional UseAsRef As Boolean) As String

	Application.Volatile
	If UseAsRef = True Then
	   SheetName = "'" & rCell.Parent.Name & "'!"
	Else
		SheetName = rCell.Parent.Name
	End If

End Function

-----------------------------------------------------------------------------------------------------------------------------------------------------

Public sub OpenFileWithFilenamePattern()
	strFilename = Dir$(ThisWorkbook.Path & "\01. Data\PS - GPE Dashboard*.xlsx")
	Do While Len(strFilename) <> 0
		Set xlBook = Workbooks.Open(Filename:=ThisWorkbook.Path & "\01. Data\" & strFilename)
		Exit Do
		strFilename = Dir$()
	Loop
	GPEDash = ActiveWorkbook.Name
End sub

-----------------------------------------------------------------------------------------------------------------------------------------------------
''Get the Last Modified File name from folder by passing Filename pattern
Function Get_LastModified_Filename(folder_Path As String, file_Pattern As String) As String
    
'''------------------------------------------------------------------------------------
'''    folder_Path = "\\teamspace.intranet.group\sites\MIRepPortal\CMT"
'''    file_Pattern = "GIT Extract_2013*.zip"
'''------------------------------------------------------------------------------------
    
    Dim strFileName As String
    Dim finalFileName As String
    Dim date_Value As Long
    
    Dim FSO As Scripting.FileSystemObject
    Dim Fileitem As Scripting.File
    
    Set FSO = New Scripting.FileSystemObject

    strFileName = Dir$(folder_Path & "\" & file_Pattern)
    date_Value = 0
    
    Do While Len(strFileName) <> 0
       Set Fileitem = FSO.GetFile(folder_Path & "\" & strFileName)
        If date_Value < Fileitem.DateLastModified Then
            date_Value = Fileitem.DateLastModified
            finalFileName = strFileName
        End If
        strFileName = Dir$()
    Loop
    Get_LastModified_Filename = finalFileName
    
End Function

-----------------------------------------------------------------------------------------------------------------------------------------------------
''Worksheet Change- Identify when a cell value change
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim WatchRange As Range
    Dim IntersectRange As Range
    
    Set WatchRange = Range("D3:D200")
    Set IntersectRange = Intersect(Target, WatchRange)
    
    If IntersectRange Is Nothing Then
    Else
        
        MsgBox ("Cell value changed---Do some changes")
    End If

End Sub

-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub RangeTableToArray()

''VBA to move Table data to array
    Dim invWS As Worksheet
    Dim InvItems As Variant
    
    Set invWS = ThisWorkbook.ActiveSheet
    InvItems = invWS.ListObjects("InvoiceItems").DataBodyRange.Value

''VBA to move data from Range to array
	Dim DeleteEventArray as Variant
    Dim rngEventNames As Range
    Set rngEventNames = ExtractWS.Range("tblDeleteEvents")
    DeleteEventArray = Split(Join(Application.Transpose(rngEventNames.Value)))

End Sub


-----------------------------------------------------------------------------------------------------------------------------------------------------
''VBA code to Filter and copy to other sheet
Public sub FilterAndCopyToAnotherSheet()

	Set RIExtract_WB = Application.Workbooks.Open(File_path & "\" & Range("Risk_Extract_File"), , True)
	Set RIExtract_WS = RIExtract_WB.Sheets("DT_Risks")

	RIExtract_WS.AutoFilterMode = False
	Cells.AutoFilter field:=1, Criteria1:=PRN_Num, Operator:=xlFilterValues
	RIExtract_WS.AutoFilter.Range.Copy

	Temp_Risk_WS.Range("A1").PasteSpecial xlPasteValues
	Application.CutCopyMode = False

	RIExtract_WB.Close False

End sub

-----------------------------------------------------------------------------------------------------------------------------------------------------
''Filter and Delete the Filtered rows 
Public sub FilterAndDeleteTheFilteredRows()

	On Error Resume Next
	CurrBENFC_WS.AutoFilterMode = False
	Cells.AutoFilter Field:=3, Criteria1:="=0", Operator:=xlFilterValues
	ActiveSheet.UsedRange.Offset(1, 0).Resize(ActiveSheet.UsedRange.Rows.Count - 1).Rows.Delete

	CurrBENFC_WS.AutoFilterMode = False
	Range("A1").Select
	On Error Goto 0

'(OR)
	ActiveSheet.UsedRange. Range("A2:A" & outputRwCnt).SpecialCells(xlCellTypeVisible).EntireRow.Delete
End sub


''VBA Function to Filter and Delete Non filtered Rows
Function Filter_And_Delete_Others(WB_File As Workbook, WS_Name As String, Filter_Column As Integer, Filter_values As Variant) As Integer

	Dim Temp_Sheet_Name As String
	Dim MIRP_WS As Worksheet

	On Error GoTo ErrorHandler

	Set MIRP_WS = WB_File.Sheets(WS_Name)

	MIRP_WS.Activate
	Temp_Sheet_Name = MIRP_WS.Name
	MIRP_WS.Name = "To_Delete"

	MIRP_WS.AutoFilterMode = False
	Cells.AutoFilter field:=Filter_Column, Criteria1:=Filter_values, Operator:=xlFilterValues
	MIRP_WS.AutoFilter.Range.Copy
	Sheets.Add After:=Sheets(Sheets.Count)
	ActiveSheet.Paste
	Range("A1").Select
	ActiveSheet.Name = Temp_Sheet_Name

	MIRP_WS.Delete
	Filter_And_Delete_Others = 0
	Exit Function

	ErrorHandler:
		Filter_And_Delete_Others = -1
    
End Function

-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub NumberOfRowsInAutoFilter()

	''VBA code to find the number of rows in the Auto filter
	Filter_Rows = WS.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count – 1

End sub

-----------------------------------------------------------------------------------------------------------------------------------------------------
''Save excel to and Excel file
Public Sub SaveToExcel(wrkSheet As Worksheet, fileName As String)

    Dim fldrName As String
    Dim fName As String
    
    fName = fileName & "_" & Format(Now(), "MMMDDYYYY_HHMM") & ".xlsx"
    
    '---- Below code will display the FolderDialog window for user to select the folder ------
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.count > 0 Then
            fldrName = .SelectedItems(1)
        End If
    End With

    If Len(fldrName) > 0 Then
        wrkSheet.Select
        wrkSheet.Copy
                
        ActiveWorkbook.SaveAs fileName:=fldrName & "\" & fName, _
            FileFormat:=51, Password:="", WriteResPassword:="", _
            ReadOnlyRecommended:=False, CreateBackup:=False
        
        ActiveWindow.Close
    End If
	'51 = xlOpenXMLWorkbook (without macro's in 2007-2010, xlsx)
	'52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2010, xlsm)
	'50 = xlExcel12 (Excel Binary Workbook in 2007-2010 with or without macro's, xlsb)
	'56 = xlExcel8 (97-2003 format in Excel 2007-2010, xls)
End Sub

-----------------------------------------------------------------------------------------------------------------------------------------------------
''Open file dialogue box 
Public Sub OpenFileDialogueBox()
	Dim Data_File_Name as variant
	Data_File_Name = Application.GetOpenFilename(FileFilter:="Microsoft Excel Files,*.xls;*.xlsx", Title:="Please select a file")

	If Data_File_Name <> False Then
		Workbooks.Open Filename:=Data_File_Name
		Set BENData_WB = Application.ActiveWorkbook
		Set BENData_WS = Application.ActiveSheet
	End if
End sub

-----------------------------------------------------------------------------------------------------------------------------------------------------
''VBA code to convert Column to Letter
Function ConvertToLetter(iCol As Integer) As String
   ConvertToLetter = Split(Cells(1, iCol).Address, "$")(1)
End Function

-----------------------------------------------------------------------------------------------------------------------------------------------------
''VBA code to combine multiple .xls files in a directory to single .xls file with multiple sheets
Public Sub Combine() 
         
    Fpath = "C:Temp" ' change to suit your directory
    Fname = Dir(FilePth & "*.xls") 
     
    Do While Fname <> "" 
        Workbooks.Open Fpath & Fname 
        Sheets(1).Copy After:=Workbooks("Master.xls").Sheets(Workbooks("Master.xls").Sheets.Count) 
        Workbooks(Fname).Close SaveChanges:=False 
        Fname = Dir 
    Loop 
    
End Sub

-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub hideAllWorksheets()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 1) <> "_" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

End Sub

-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub DisplayAllWorksheets()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws

End Sub

-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub getQueryResultsToExcel() 'AccessFilePath as string, AccessPassword as String, SQLStr As String, outputRange As Range)

    Dim cn As Object
    Dim rs As Object
    Dim strFile As String
    Dim strCon As String
    Dim strPassword As String

    strFile = "C:\MANOHAR\PQQ Offline MI_Acc v0.5_14Dec v0.1.accdb"      ''Access database file path
    strPassword = "payQwiQ"
	
	If Len(strPassword) > 0 then
		strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile & ";Jet OLEDB:Database Password=" & strPassword & ";"
	Else
		strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile & ";"
    End if

    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    cn.Open strCon

    SQLStr = "SELECT * FROM VBA_PQQ_Totals"
    
    '----- Run the query ----------
    rs.Open SQLStr, cn, 3, 3
    
    '------ Move the date to the sheet ------
    outputRange.CopyFromRecordset rs
    
    '------ Close and return the connection and recordset ------
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

End Sub


Public Sub accessToExcel() 'SQLStr As String, Optional queryName As String)

    Dim cn As Object
    Dim rs As Object
    Dim strFile As String
    Dim strCon As String
    Dim strPassword As String

    strFile = "C:\MANOHAR\PQQ Offline MI_Acc v0.5_10Dec v0.3.accdb"      ''Access database file path
    strPassword = "testPassword"
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile & ";Jet OLEDB:Database Password=" & strPassword & ";"
    'strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile & ";"
    
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    cn.Open strCon

    SQLStr = "SELECT * FROM AllData WHERE ID = 12"
    
    '----- Run the query ----------
    rs.Open SQLStr, cn, 3, 3
    
    '------ Move the date to the sheet ------
    Worksheets("Sheet1").Cells(2, 1).CopyFromRecordset rs
    
    '------ Close and return the connection and recordset ------
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

End Sub

''-----------------------------------------------------------------------------------------------------------------------------------------------------

Public Function OpenAccessToExcelConnection(ConnObject As Object, accessFileName As String, Optional accessPwd As String) As Integer

    Dim strCon As String

    On Error GoTo ErrorHandler

    If Len(accessPwd) > 0 Then
        strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessFileName & ";Jet OLEDB:Database Password=" & accessPwd & ";"
    Else
        strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessFileName & ";"
    End If

    ConnObject.Open strCon
    
    OpenAccessToExcelConnection = 0
    Exit Function
    
ErrorHandler:
    OpenAccessToExcelConnection = 1
    
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------

Public Function getAccessQuerytoExcelRange(OLEDBConn As Object, SQLStr As String, opRange As Range) As Integer

On Error GoTo ErrHandler

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    '----- Run the query ----------
    rs.Open SQLStr, OLEDBConn, 3, 3
    
    '------ Move the date to the sheet ------
    opRange.CopyFromRecordset rs
    
    '------ Close and return the recordset ------
    rs.Close
    Set rs = Nothing
    getAccessQuerytoExcelRange = 0
    Exit Function
    
ErrHandler:
    Set rs = Nothing
    getAccessQuerytoExcelRange = 1

End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------

Public Function CloseAccessToExcelConnection(Conn As Object) as Integer
	On Error GoTo ErrHandler
		Conn.Close
		Set Conn = Nothing
		
		CloseAccessToExcelConnection = 0
		Exit Function
		
	ErrHandler:
		CloseAccessToExcelConnection = 1

End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------

Sub testdb()

    Dim con As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim rs As ADODB.Recordset

    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command

    With con
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open "H:\WBC\Lukas\STOP.accdb"
    End With

    With cmd
        .ActiveConnection = con
        .CommandText = "Query1"
        .CommandType = adCmdStoredProc

        .Parameters.Append cmd.CreateParameter("MyID", adInteger, adParamInput)
        .Parameters("MyID") = 1
    End With

    Set rs = New ADODB.Recordset
    rs.Open cmd

    Do Until rs.EOF
        Debug.Print rs.Fields("ID").Value
        rs.MoveNext
    Loop

    rs.Close
    con.Close

    Set cmd = Nothing
    Set rs = Nothing
    Set prm = Nothing
    Set con = Nothing

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
''VBA SQL Routine sample 

Sub Populate_Top5_Risks(sPRN As String)

'    Dim sPRN As String
'    sPRN = "GY132"
    
    Dim szConnect As String
    Dim SQL As String
    Dim rsData As ADODB.Recordset
    
    Dim risksWS, prjDashboardWS As Worksheet
    Set risksWS = ThisWorkbook.Sheets("Risks")
    Set prjDashboardWS = ThisWorkbook.Sheets("Project Dashboard")

    Dim top5RisksRng As String
    top5RisksRng = Range("Top5Risks").Cells(1, 1).Address
    
    
    '--- Set Name range to access via the SQL query ----
    risksWS.AutoFilterMode = False
    ThisWorkbook.Names("SQLTBL_RISKS").RefersTo = "=Risks!$A$1:$AF$1000"
    
    '--- Set up the connection string to excel
    szConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=Excel 8.0;"
    
    
    '-- Extract and populate Top 5 risks ------------
    Set rsData = New ADODB.Recordset
    SQL = "SELECT [Log Ref], [Risk Title],'','','','','','','','','', [Impact],[Probability],[Impact Date], [Next Review Date]  from SQLTBL_RISKS where [PRN] = '" & sPRN & "'" & _
    " Order by [Ranking], [Impact Date]"
    
    '-- Run the query as adCmdText and if data found past them in the Top 5 range
    rsData.Open SQL, szConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
    prjDashboardWS.Range("Top5Risks").ClearContents
    If Not rsData.EOF Then
        prjDashboardWS.Range(top5RisksRng).CopyFromRecordset rsData, 5
    End If
    rsData.Close
    Set rsData = Nothing
    
    
    '------------------------------------------------------------
    '----- Extract and populate Non-Compliant Risks  ------------
    '------------------------------------------------------------
    Set rsData = New ADODB.Recordset
    prjDashboardWS.Range("NonCompRisks").Value = ""
    
    SQL = "SELECT [Log Ref] from SQLTBL_RISKS where [PRN] = '" & sPRN & "' and [Non-Compliant]=TRUE"
    rsData.Open SQL, szConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Dim nonCmpRisks As String
    Do While Not rsData.EOF
       nonCmpRisks = nonCmpRisks & rsData.Fields("Log Ref") & ", "
       rsData.MoveNext
    Loop
    
    If Len(nonCmpRisks) > 0 Then
        prjDashboardWS.Range("NonCompRisks").Value = Left(nonCmpRisks, Len(nonCmpRisks) - 2)
    End If
    
    rsData.Close
    Set rsData = Nothing

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
''VBA Inputbox code to get a value from the user
Public Sub InputBoxToGetValueFromUser()
	Dim response As Variant 
	 
	response = InputBox("Prompt", "Title") 
	Select Case StrPtr(response) 
	Case 0 
		 'OK not pressed
		Exit Sub 
	Case Else 
		 'OK pressed
		 'Carry on your routine, variable response contains the InputText
	End Select
End sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
''Filter by Date or Date filter
Sub FilterByExactDate()
	Dim dDate As Date
	Dim strDate As String
	Dim lDate As Long
		dDate = DateSerial(2006, 8, 12)
		lDate = dDate
		Range("A1").AutoFilter
		Range("A1").AutoFilter Field:=1, Criteria1:=">=" & lDate, _
						 Operator:=xlAnd, Criteria2:="<" & lDate + 1
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
''Filter by Exact date & Time
Sub FilterByDateTime()
	Dim dDate As Date
	Dim dbDate As Double
	If IsDate(Range("B1")) Then
		dbDate = Range("B1")
		dbDate = DateSerial(Year(dbDate), Month(dbDate), Day(dbDate)) + _
			 TimeSerial(Hour(dbDate), Minute(dbDate), Second(dbDate))
		Range("A1").AutoFilter
		Range("A1").AutoFilter Field:=1, Criteria1:=">" & dbDate
	End If
End Sub	

'-----------------------------------------------------------------------------------------------------------------------------------------------------
''Save string to Array based on the delimiter character
Public Sub SaveStringToArrayBasedOnDelimiter()
    Dim WBS_Code_Text As String
    Dim WBS_Code_Array As Variant
    WBS_Code_Text = Range("WBS_Codes").Text
    WBS_Code_Array = Split(WBS_Code_Text, "/") 'create array of values using "/" as delimiter
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExcelFormulas()
''''Formula to get the Last word in a Cell
'=TRIM(RIGHT(SUBSTITUTE(TRIM(A1)," ",REPT(" ",99)),99))


''''Formula to get the First word in a cell
'=TRIM(LEFT(SUBSTITUTE(TRIM(A1)," ",REPT(" ",99)),99))


''Formula to get the nth Occurrence of the Character

''''Below will find the 2nd occurrence of  “|” in cell I12
'=FIND("|",I12,FIND("|",I12)+1)

''''Below will find the 3rd Occurrence of ‘c’ in the cell A1
'=FIND(CHAR(1),SUBSTITUTE(A1,"c",CHAR(1),3))


End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExtractJSONValue(strFieldName As String, strJSON As String, Optional StrFieldName2 As String, Optional outputAsNumeric As Boolean, Optional eventName As String)

    Dim strWrkStr As String
    Dim LenStrFieldNm As Long, LenStrJSON As Long, PosStrFldNm As Long
    
    strWrkStr = ""
    
    On Error GoTo Errhandler
    
    '----------------------- Event message based parsing -----------------------------
    
    If Len(eventName) > 0 Then
        '---- For LOGIN events request message will contain only Email ----
        If eventName = "LOGIN" And StrFieldName2 = "Email" Then
            ExtractJSONValue = strJSON
            Exit Function
        End If
        
        '----- For IDENTITY_CHECK failures the response message will contian the error message ---
        If eventName = "IDENTITY_CHECK" And StrFieldName2 = "errorDescription" Then
            ExtractJSONValue = strJSON
            Exit Function
        End If

    End If
    
    '---------------------------------------------------------------------------------
    LenStrFieldNm = Len(strFieldName)
    LenStrJSON = Len(strJSON)
    PosStrFldNm = InStr(strJSON, strFieldName & Chr(34))
    
    '--- If 1st Field name search not found then search for 2nd Field name if provided in the input ---
    If PosStrFldNm = 0 Then
        If Len(StrFieldName2) > 0 Then
            LenStrFieldNm = Len(StrFieldName2)
            PosStrFldNm = InStr(strJSON, StrFieldName2 & Chr(34))
        End If
        
        If PosStrFldNm = 0 Then
            If outputAsNumeric = True Then
                ExtractJSONValue = 0
                Exit Function
            Else
                ExtractJSONValue = ""                                                       'ExtractJSONValue = "Field not present"
                Exit Function
            End If
        End If
    End If
    
    strWrkStr = Right(strJSON, LenStrJSON - (PosStrFldNm + LenStrFieldNm + 1))
    
    Dim chr34Pos As Integer         ' Char34 is "
    Dim chr125Pos As Integer        ' Char125 is }
    
    If Left(strWrkStr, 1) = Chr(34) Then                                                '   Determines if value is a number or nullvalue
    
        chr34Pos = InStr(strWrkStr, Chr(34) & ",")
        chr125Pos = InStr(strWrkStr, Chr(34) & "}")
        
        If (chr34Pos < chr125Pos And chr34Pos > 0) Or (chr34Pos > 0 And chr125Pos = 0) Then       ' If "," delimiter comes first as delimiter
            strWrkStr = Left(strWrkStr, InStr(strWrkStr, Chr(34) & ",") - 1)
            strWrkStr = Right(strWrkStr, Len(strWrkStr) - 1)
        Else                                                                                      ' Else use "}" as the delimiter
            strWrkStr = Left(strWrkStr, InStr(strWrkStr, Chr(34) & "}") - 1)
            strWrkStr = Right(strWrkStr, Len(strWrkStr) - 1)
        End If

    Else
        If InStr(strWrkStr, ",") = 0 Then
            strWrkStr = Left(strWrkStr, InStr(strWrkStr, "}") - 1)
        Else
            strWrkStr = Left(strWrkStr, InStr(strWrkStr, ",") - 1)
        End If
    End If
    
    If outputAsNumeric = True Then
       If Len(strWrkStr) = 0 Or strWrkStr = "null" Then
          ExtractJSONValue = 0
          Exit Function
       End If
       ExtractJSONValue = strWrkStr * 1
    Else
        ExtractJSONValue = strWrkStr
    End If
    Exit Function
    
    '---- Error Handler-----------
Errhandler:
    If outputAsNumeric = True Then
       ExtractJSONValue = 0
    Else
       ExtractJSONValue = ""
    End If

End Function



'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExtractXMLValue(strFieldName As String, strXML As String, Optional strFieldName2 As String, Optional outputAsNumeric As Boolean)
    
    Dim strWrk As String
    Dim LenStrFieldNm As Long, LenStrXML As Long, startPos As Long, endPos As Long, dataLen As Integer
    
    strWrk = ""
    
    On Error GoTo ErrHandler
    
    '---- 1. Search if first field exists in the xml data -----
    LenStrXML = Len(strXML)
    LenStrFieldNm = Len(strFieldName)
    startPos = InStr(strXML, "<" & strFieldName & ">")
    endPos = InStr(strXML, "</" & strFieldName & ">")
    '---- 2. If first field does not exist check if 2nd field exists in the xml data----
    If startPos = 0 Then
        If Len(strFieldName2) > 0 Then
            LenStrFieldNm = Len(strFieldName2)
            startPos = InStr(strXML, "<" & strFieldName2 & ">")
            endPos = InStr(strXML, "</" & strFieldName2 & ">")
        End If
        
        If startPos = 0 Then                 '-- Both fields not present
            If outputAsNumeric = True Then
                ExtractXMLValue = 0
                Exit Function
            Else
                ExtractXMLValue = ""
                Exit Function
            End If
        End If
    End If

    '----3. Find the length of the XML field data --------
    startPos = startPos + LenStrFieldNm + 2
    dataLen = endPos - startPos
    strWrk = Mid(strXML, startPos, dataLen)
    
    If outputAsNumeric = True Then
       If Len(strWrk) = 0 Or strWrk = "null" Then
          ExtractXMLValue = 0
       Else
          ExtractXMLValue = strWrk * 1
       End If
    Else
        ExtractXMLValue = strWrk
    End If
    
    Exit Function

ErrHandler:

    If outputAsNumeric = True Then
       ExtractXMLValue = 0
    Else
       ExtractXMLValue = ""
    End If
    
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function CheckFileExists(strFilePathName As String) As Boolean

    'Dim strFilePathName As String
    'strFilePathName = "C:\Temp\Test File.xlsx"
    Dim strFileExists As String
    strFileExists = Dir(strFilePathName)

    If strFileExists = "" Then
         CheckFileExists = False
     Else
         CheckFileExists = True
     End If

End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function CheckFolderExists(strFolderPathName As String) As Boolean

'    Dim strFolderPathName As String
'    strFolderPathName = "C:\Temp\"
    
    Dim strFolderExists As String
    strFolderExists = Dir(strFolderPathName, vbDirectory)
    
    If strFolderExists = "" Then
        CheckFolderExists = False
    Else
        CheckFolderExists = True
    End If

End Function
