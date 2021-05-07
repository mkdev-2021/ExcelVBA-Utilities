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

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
    
End Function
