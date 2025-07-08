Attribute VB_Name = "Module3"
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Sub Copy_Sheet()
MsgBox "Select import file"
FileLocation = Application.GetOpenFilename
If FileLocation = "False" Then
    Beep
    Exit Sub
End If

Set wb = ActiveWorkbook
Set sl = wb.Worksheets("sheet_list")

Application.ScreenUpdating = False
Set ImportWorkbook = Workbooks.Open(Filename:=FileLocation)

row_st = 2
row_en = wb.Worksheets("sheet_list").Cells(Rows.Count, "A").End(xlUp).Row

For i = row_st To row_en
    ws_new = wb.Worksheets("sheet_list").Cells(i, "A").Value
    ws_old = wb.Worksheets("sheet_list").Cells(i, "B").Value
        
    If WorksheetExists(CStr(ws_old)) Then
        Set wsPaste = wb.Worksheets(ws_new)
        Set wsImport = ImportWorkbook.Worksheets(ws_old)
        wsImport.Range("A9:D50").Copy
        wsPaste.Range("A9").PasteSpecial xlPasteValues
    End If
Next i
    
ImportWorkbook.Close False

End Sub
