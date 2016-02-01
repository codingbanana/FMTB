Attribute VB_Name = "Regulator_Assessment"
Sub findMeasID()
'Hao Zhang @ 2015.7.6
'create a hash table between RegID and MeasID

Set tgtws = ThisWorkbook.Sheets("SiteMeasRel")
Set srcws = ThisWorkbook.Sheets("MeasData")

For i = 2 To 206
    k = 0
    For j = 269 To 2 Step -1
        If srcws.Cells(j, 3).Value = tgtws.Cells(i, 1).Value Then
            If srcws.Cells(j, 1).Value <> tgtws.Cells(i, 2).Value Then
                tgtws.Cells(i, k + 3).Value = srcws.Cells(j, 1).Value
                k = k + 1
            End If
        End If
    Next
Next

End Sub

