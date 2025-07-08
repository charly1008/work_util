Attribute VB_Name = "Module2"
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Sub Archive()
Attribute Archive.VB_ProcData.VB_Invoke_Func = " \n14"
'Remove redundant comment rows
With ActiveWorkbook
    For Each ws In Worksheets
        If Not IsInArray(ws.Name, Array("master", "template")) Then
            ws.Activate
            With ActiveSheet
                Row = ws.Cells(Rows.Count, "C").End(xlUp).Row
                
                'Clear borders
                Range("A9:D20").Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                Selection.Borders(xlEdgeLeft).LineStyle = xlNone
                Selection.Borders(xlEdgeTop).LineStyle = xlNone
                Selection.Borders(xlEdgeBottom).LineStyle = xlNone
                Selection.Borders(xlEdgeRight).LineStyle = xlNone
                Selection.Borders(xlInsideVertical).LineStyle = xlNone
                Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                     
                'Add borders
                Range("A9:D" & Row).Select
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Range("A1").Select
            End With
        End If
    Next ws
End With

'Check hyperlinks
With ActiveWorkbook.Worksheets("master")
    row_st = .Columns("A").Find("Category").Row
    row_en = .Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = row_st To row_en
        Name = .Cells(i, "A").Value
        
        If InStr(1, Name, "Category") = 0 And Not Name = "" Then
            If .Cells(i, "A").Hyperlinks.Count > 0 Then
                Link0 = .Cells(i, "A").Hyperlinks(1).SubAddress
                Link = Mid(Link0, 2, InStr(1, Link0, "!") - 3)
                If Name = Link Then
                    .Cells(i, "L") = ""
                Else
                    .Cells(i, "L") = Link
                End If
            Else
                .Cells(i, "L") = "Missing"
            End If
            
            If .Cells(i, "L") = "" Then
                .Cells(i, "L").Interior.Color = RGB(255, 255, 255)
            Else
                .Cells(i, "L").Interior.Color = RGB(255, 255, 0)
            End If
        End If
    Next i
End With

'Check comment response
With ActiveWorkbook
    For Each ws In Worksheets
        If Not IsInArray(ws.Name, Array("master", "template")) Then
            ws.Activate
            With ActiveSheet
                Row = ws.Cells(Rows.Count, "C").End(xlUp).Row
                For i = 9 To Row
                    If .Cells(i, "B") <> "" And .Cells(i, "C") <> "" And .Cells(i, "D") <> "" Then
                        .Cells(i, "E") = ""
                    Else
                        .Cells(i, "E") = "Check this row!"
                    End If
                Next i
                Range("A1").Select
            End With
        End If
    Next ws
End With

End Sub
