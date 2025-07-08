Attribute VB_Name = "Module1"
Sub Create_Sheet()
With ActiveWorkbook
    row_st = .Worksheets("master").Columns("A").Find("Category").Row
    row_en = .Worksheets("master").Cells(Rows.Count, "A").End(xlUp).Row
    
    Sheet = ""
    For Each ws In Worksheets
        Sheet = Sheet + """" + ws.Name + ""","
    Next ws
    
    For i = row_st To row_en
        With .Worksheets("master")
            Name = .Cells(i, "A").Value
            
            If InStr(1, Name, "Category") Then
                If InStr(1, Name, "SDTM") Then
                    CATE = "SDTM\program"
                ElseIf InStr(1, Name, "ADaM") Then
                    CATE = "ADaM\program"
                Else
                    CATE = "Tables"
                End If
            Else
                If .Cells(i, "A") <> "" Then
                    .Cells(i, "M") = "=IF(COUNTIF(INDIRECT(""'""&$A" + CStr(i) + "&""'!E9:E200""),""<>"")=0,"""",""Need to check"")"
                End If
            End If

            
            Namex = "=master!$A$" + CStr(i)
            Prim = "=master!$D$" + CStr(i)
            Vali = "=master!$F$" + CStr(i)
            If InStr(1, UCase(Name), "SPEC") = 0 Then
                Prim_pg = "=CONCATENATE(master!$B$3,""\" + CATE + "\Production\"",master!$E$" + CStr(i) + ","".sas"")"
                Vali_pg = "=CONCATENATE(master!$B$3,""\" + CATE + "\Validation\"",master!$G$" + CStr(i) + ","".sas"")"
            Else
                'Prim_pg = "=CONCATENATE(master!$B$3,""\" + CATE + "\doc\Mapping_define\"",master!$E$" + CStr(i) + ","".xls"")"
                Prim_pg = ""
                Vali_pg = ""
            End If
        End With
        
        If InStr(1, Name, "Category") = 0 And Not Name = "" And InStr(1, Sheet, """" + Name + """") = 0 Then
            Worksheets("template").Copy After:=.Worksheets(.Worksheets.Count)
            ActiveSheet.Name = Name

            With ActiveSheet
                .Cells(2, "B") = Namex
                .Cells(3, "B") = Prim
                .Cells(4, "B") = Prim_pg
                .Cells(5, "B") = Vali
                .Cells(6, "B") = Vali_pg
            End With
            
            With Worksheets("master")
                .Hyperlinks.Add Anchor:=.Cells(i, "A"), Address:="", SubAddress:="'" + Name + "'!$A$1", TextToDisplay:=Name
            End With
        End If
    Next i
End With

End Sub
