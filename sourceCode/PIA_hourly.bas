Attribute VB_Name = "PIA_hourly"
Sub dateTimeConvert()
'Hao Zhang @ 2015.2.16

'for PIA hourly data update, PIA_MMMYYYY.xlsx
'convert Date and time column with proper format

Dim wb As Workbook
Dim ws As Worksheet
Set wb = ActiveWorkbook
Set ws = wb.ActiveSheet
    
    
    CurrentYear = Mid(wb.Name, 8, 4)
    'CurrentMonth = MATCH(mid(wb.Name,5,3),{"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"},0)
    currentMonth = Month(1 & " " & Mid(wb.Name, 5, 3))
    
For iRow = 6 To Cells(6, "B").End(xlDown).Row
    Cells(iRow, 1).Value = DateSerial(CurrentYear, currentMonth, Cells(iRow, 1).Value)
    p = Cells(iRow, 2).Value
    Cells(iRow, 2).Value = Format(TimeSerial(WorksheetFunction.Floor(p / 100, 1), Right(p, 2), 0), "hh:mm")
Next

End Sub
