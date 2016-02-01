Attribute VB_Name = "RegDataExtract"

Private Sub revis()
'Hao Zhang @ 20150610
'fill values in site_Revisited and Times of Revisit cols

Dim ws As Worksheet
Set ws = ActiveWorkbook.ActiveSheet
With ws
    EOR = ws.UsedRange.Rows.count
    For i = 2 To EOR
        For j = i + 1 To EOR
            If .Cells(i, 4).Value = .Cells(j, 4).Value Then
                If .Cells(j, 2).Value > .Cells(i, 2).Value Then
                    If .Cells(i, 19).Value = .Cells(j, 19).Value Then
                        .Cells(i, 8).Value = "Yes"
                        .Cells(i, 9).Value = 1
                        .Cells(j, 9).Value = 2
                    End If
                End If
            End If
        Next
        .Cells(i, 8).Value = "No"
        .Cells(i, 9).Value = 1
    Next
End With
            

End Sub

Private Sub addlink()
'Hao Zhang @ 2015.7.1
'Add link to videos


Dim i As Integer, j As Integer

With Workbooks("Regulator_Assessment_data_post-processing_v5.xlsx")
    Set ws1 = .Sheets("SiteInfo")
    Set ws2 = .Sheets("Video")
End With

For i = 2 To ws1.UsedRange.Rows.count
    For j = 9 To 10
        If ws1.Cells(i, j) <> "" Then
        ws1.Hyperlinks.Add anchor:=ws1.Cells(i, j), Address:=ws2.Columns(1).Find(ws1.Cells(i, j).Value, lookat:=xlWhole).Offset(0, 1).Value
        End If
    Next
Next

For i = 2 To ws2.UsedRange.Rows.count
    For j = 5 To 6
        If ws2.Cells(i, j) <> "" Then
        ws2.Hyperlinks.Add anchor:=ws2.Cells(i, j), Address:=ws2.Columns(1).Find(ws2.Cells(i, j).Value, lookat:=xlWhole).Offset(0, 1).Value
        End If
    Next
Next

End Sub
Private Sub addlink2()
'Hao Zhang @ 2015.7.1
'Add link to FieldReport

Set ws = Workbooks("Regulator_Assessment_data_post-processing_v5.xlsx").Sheets("FieldReport")

For i = 2 To ws.UsedRange.Rows.count
        ws.Hyperlinks.Add anchor:=ws.Cells(i, 1), Address:=ws.Cells(i, 2).Value
Next


End Sub

Private Sub addlink3()
'Hao Zhang @ 2015.7.10
'Add link to ModelValidation

Set ws = Workbooks("Regulator_Assessment_data_post-processing_v5.xlsx").Sheets("ModelValidation")

For i = 2 To 206
    If IsError(ws.Cells(i, 31).Value) = False Then
        ws.Hyperlinks.Add anchor:=ws.Cells(i, 31), Address:=ws.Cells(i, 31).Value
        ws.Cells(i, 31).Value = ws.Cells(i, 2).Value
    Else
        ws.Cells(i, 31).Value = ""
    End If
Next

End Sub
Private Sub addlink4()
'Hao Zhang @ 2015.7.10
'Add link to ModelValidation

Set ws = Workbooks("Regulator_Assessment_data_post-processing_v5.xlsx").Sheets("ModelValidation")

For i = 2 To 206
    If ws.Cells(i, 30).Value <> "" Then
        ws.Hyperlinks.Add anchor:=ws.Cells(i, 29), Address:=ws.Cells(i, 30).Value
    End If
    If ws.Cells(i, 28).Value <> "" Then
        ws.Hyperlinks.Add anchor:=ws.Cells(i, 27), Address:=ws.Cells(i, 28).Value
    End If
Next

End Sub
