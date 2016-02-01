Attribute VB_Name = "populate_time_range"
Sub test()
'Hao Zhang @ 2015.2.1
'tested the performance for 4 methods of populating a time range in a worksheet
Row1 = 14
Row4 = 80000
month1 = DateSerial(2015, 4, 1)
intvl = 15
t0 = Timer()
'1. the loop method: slowest
With ActiveWorkbook.Worksheets(1)
'clear contents of previous dates
    .range("A" & Row1, .Cells(.Rows.count, "A").End(xlUp)).ClearContents
'add the first date of the quarter
    .range("A" & Row1).Value = month1
'not needed as values are directly filled
''If .Range("A" & Row4).Formula <> "=A" & Row4 - 1 & "+(" & intvl & "/60)/24" Then
    For iRow = Row1 + 1 To Row4
    'add formulas to the next rows
    ''.Range("A" & iRow) = "=A" & iRow - 1 & "+(" & intvl & "/60)/24"
    'add values to the next rows (faster)
    .Cells(iRow, "A").Value = Format(DateAdd("n", intvl * (iRow - Row1), month1), "mm/dd/yyyy hh:mm:ss")
    DoEvents
    Next
    .range("A" & Row1, "A" & Row4).NumberFormat = "mm/dd/yyyy hh:mm:ss"
'End If
End With
t1 = Timer() - t0
'2. the array method: faster than first method
Dim arr() As Date
'1-D array is multi COLUMN, not Row!!!
'for 1-D array, use application.worksheetsfunction.transpose() to convert columns into rows
'however, the transpose can only handle 2^16=64000 data
'Therefore, it's better to dim a 2D array with single column
ReDim arr(Row1 To Row4, 1 To 1)
arr(Row1, 1) = Format(month1, "mm/dd/yyyy hh:mm:ss")
For iRow = Row1 + 1 To Row4
    arr(iRow, 1) = Format(DateAdd("n", intvl * (iRow - Row1), month1), "mm/dd/yyyy hh:mm:ss")
    DoEvents
Next
With ActiveWorkbook.Worksheets(2)
    'clear contents of previous dates
    .range("A" & Row1, .Cells(.Rows.count, "A").End(xlUp)).ClearContents
    .range("A" & Row1, "A" & Row4).Value = arr
    .range("A" & Row1, "A" & Row4).NumberFormat = "mm/dd/yyyy hh:mm:ss"

'ActiveWorkbook.Worksheets(2).Range("A" & Row1).Resize(UBound(Arr) - LBound(Arr) + 1, 1).Value = Arr
End With
T2 = Timer() - t0 - t1

'3. the autofill method: faster than method 1 and 2
With ActiveWorkbook.Worksheets(3)
'clear contents of previous dates
    .range("A" & Row1, .Cells(.Rows.count, "A").End(xlUp)).ClearContents
'add the first date of the quarter
    .range("A" & Row1).Value = month1
    .range("A" & Row1 + 1).Value = DateAdd("n", intvl, month1)
    .range("A" & Row1 + 2).Value = DateAdd("n", intvl * 2, month1)
    Set SourceRange = .range("A" & Row1, "A" & Row1 + 2)
    Set fillRange = .range("A" & Row1 + 2, "A" & Row4)
    SourceRange.AutoFill Destination:=fillRange
    .range("A" & Row1, "A" & Row4).NumberFormat = "mm/dd/yyyy hh:mm:ss"
End With
T3 = Timer - t0 - t1 - T2
'4. fill formula method: the fastest
With ActiveWorkbook.Worksheets(4)
'clear contents of previous dates
    .range("A" & Row1, .Cells(.Rows.count, "A").End(xlUp)).ClearContents
    
    .range("A" & Row1).Value = month1
    .range("A" & Row1 + 1, "A" & Row4).Formula = "=A" & Row1 & "+(" & intvl & "/60)/24"
    .range("A" & Row1, "A" & Row4).NumberFormat = "mm/dd/yyyy hh:mm:ss"
End With
t4 = Timer() - t0 - t1 - T2 - T3
MsgBox "loop method: " & t1 & Chr(13) & "array method: " & T2 & Chr(13) & "fill loop method: " & T3 & Chr(13) & "fill method: " & t4

End Sub

