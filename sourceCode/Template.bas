Attribute VB_Name = "Template"
Sub changeScale()
'Hao Zhang @ 2015.1.13
'update QA template to new quarter
'Base template: Q4-2014.xlsm 12/31/2014

'********set the year and first month of the new quarter here********
mYear = 2015
startMonth = 1
'enter the time interval in minutes
intvl = 2
'********************************************************************

'need to add:1. tab name update, 2. month update for tables


startDate = DateSerial(mYear, startMonth, 1)
endDate = DateAdd("m", 1, startDate)

'identify the row number of the begining of each month (Row1,2,3) and the end of the last month(Row4)
    'For Each cell In .Range("A14", "A" & i)
    'If cell.Value = startDate Then
    '    Row1 = cell.Row
    'End If
    'If cell.Value = DateAdd("m", 1, startDate) Then
    '    Row2 = cell.Row
    'End If
    'If cell.Value = DateAdd("m", 2, startDate) Then
    '    Row3 = cell.Row
    'End If
    'Next
    'Row4 = i

'identify the row number of the begining of each month (Row1,2,3) and the end of the last month(Row4)
Row1 = 14
Row2 = Row1 + ((Day(DateSerial(mYear, startMonth + 1, 1) - 1)) * 60 / intvl * 24)
Row3 = Row2 + ((Day(DateSerial(mYear, startMonth + 2, 1) - 1)) * 60 / intvl * 24)
Row4 = Row3 + ((Day(DateSerial(mYear, startMonth + 3, 1) - 1)) * 60 / intvl * 24) - 1


With ActiveWorkbook.Worksheets("Flow Data")
'clear contents of previous dates
    '.Range("A14", Cells(Rows.Count, "A").End(xlUp)).ClearContents
'add the first date of the quarter
    '.Range("A14").Value = startDate
'add formulas to the next rows
    'i = 15
    'Do While .Cells(i - 1, 1).Value < (DateAdd("m", 3, startDate))
    '.Cells(i, 1).Value = "=A" & (i - 1) & "+(" & Intvl & "/60)/24"
    'i = i + 1
    'Loop
'clear contents of previous dates
    .range("A" & Row1, .Cells(Rows.count, "A").End(xlUp)).ClearContents
'add the first date of the quarter
    .range("A" & Row1).Value = startDate
'add formulas to the next rows
If range("A" & Row4).Formula <> "=A" & Row4 - 1 & "+(" & intvl & "/60)/24" Then
    For iRow = Row1 + 1 To Row4
    .range("A" & iRow) = "=A" & iRow - 1 & "+(" & intvl & "/60)/24"
    DoEvents
    Next
End If

'change the Percent Recovery in Flow Data
    .range("I5").Value = "=(COUNT(U" & Row1 & ":U" & (Row2 - 1) & "))/(COUNT(A" & Row1 & ":A" & (Row2 - 1) & "))"
    .range("I6").Value = "=(COUNT(U" & Row2 & ":U" & (Row3 - 1) & "))/(COUNT(A" & Row2 & ":A" & (Row3 - 1) & "))"
    .range("I7").Value = "=(COUNT(U" & Row3 & ":U" & Row4 & "))/(COUNT(A" & Row3 & ":A" & Row4 & "))"
    .range("j5").Value = "=(COUNT(w" & Row1 & ":w" & (Row2 - 1) & "))/(COUNT(A" & Row1 & ":A" & (Row2 - 1) & "))"
    .range("j6").Value = "=(COUNT(w" & Row2 & ":w" & (Row3 - 1) & "))/(COUNT(A" & Row2 & ":A" & (Row3 - 1) & "))"
    .range("j7").Value = "=(COUNT(w" & Row3 & ":w" & Row4 & "))/(COUNT(A" & Row3 & ":A" & Row4 & "))"
    .range("k5").Value = "=(COUNT(v" & Row1 & ":w" & (Row2 - 1) & "))/(COUNT(A" & Row1 & ":A" & (Row2 - 1) & "))"
    .range("k6").Value = "=(COUNT(v" & Row2 & ":w" & (Row3 - 1) & "))/(COUNT(A" & Row2 & ":A" & (Row3 - 1) & "))"
    .range("k7").Value = "=(COUNT(v" & Row3 & ":w" & Row4 & "))/(COUNT(A" & Row3 & ":A" & Row4 & "))"
End With

' change the monthly charts' data range and scale
Call MonthlyChart(startDate, endDate, Row1, (Row2 - 1))
Call MonthlyChart(startDate, endDate, Row2, (Row3 - 1))
Call MonthlyChart(startDate, endDate, Row3, (Row4 - 1))

'change the scale of ALL TS chart (the data range was unchanged ending at row 8846)
Sheets("ALL TS").Select
'hydrograph
'ActiveSheet.PlotArea.Select   ''commented, hopefully it could avoid the issue that the sheet-based macro were accidentally activated
TSscaleChange (DateAdd("m", -3, startDate)), startDate
'hyetograph
ActiveSheet.ChartObjects("Rain").Activate
TSscaleChange (DateAdd("m", -3, startDate)), startDate

'all TS CORR
Sheets("ALL TS CORR").Select
'hydrograph
'ActiveSheet.PlotArea.Select
TSscaleChange (DateAdd("m", -3, startDate)), startDate
'hyetograph
ActiveSheet.ChartObjects("Rain").Activate
TSscaleChange (DateAdd("m", -3, startDate)), startDate

Sheets("ALL SP (Flow)").Select
'ActiveSheet.PlotArea.Select
'level 1
    ActiveChart.SeriesCollection("Monitored Data").XValues = "='Flow Data'!$B$" & Row1 & ":$B$" & Row4
'RAW Flow
    ActiveChart.SeriesCollection("Monitored Data").Values = "='Flow Data'!$G$" & Row1 & ":$G$" & Row4

Sheets("ALL SP CORR (Flow)").Select
'ActiveSheet.PlotArea.Select
'level 1
    ActiveChart.SeriesCollection("Monitored Data").XValues = "='Flow Data'!$B$" & Row1 & ":$B$" & Row4
'CORR Flow
    ActiveChart.SeriesCollection("Monitored Data").Values = "='Flow Data'!$W$" & Row1 & ":$W$" & Row4

Sheets("SP Flow Vs Level 1&2").Select
'ActiveSheet.PlotArea.Select
'level 1
    ActiveChart.SeriesCollection("Primary Level").XValues = "='Flow Data'!$B$" & Row1 & ":$B$" & Row4
'Flow
    ActiveChart.SeriesCollection("Primary Level").Values = "='Flow Data'!$W$" & Row1 & ":$W$" & Row4
'level 2
    ActiveChart.SeriesCollection("Redundant Level").XValues = "='Flow Data'!$D$" & Row1 & ":$D$" & Row4
'Flow
    ActiveChart.SeriesCollection("Redundant Level").Values = "='Flow Data'!$W$" & Row1 & ":$W$" & Row4

Sheets("SP Velocity Vs Level 1&2").Select
'ActiveSheet.PlotArea.Select
'level 1
    ActiveChart.SeriesCollection("Primary Level").XValues = "='Flow Data'!$B$" & Row1 & ":$B$" & Row4
'Vel 1
    ActiveChart.SeriesCollection("Primary Level").Values = "='Flow Data'!$V$" & Row1 & ":$V$" & Row4
'level 1
    ActiveChart.SeriesCollection("Redundant Level").XValues = "='Flow Data'!$D$" & Row1 & ":$D$" & Row4
'Vel 2
    ActiveChart.SeriesCollection("Redundant Level").Values = "='Flow Data'!$V$" & Row1 & ":$V$" & Row4

Sheets("SP Raw Flow Vs Corr Flow").Select
'ActiveSheet.PlotArea.Select
'RAW Flow
    ActiveChart.SeriesCollection("Calc Flow Vs Raw FLow").XValues = "='Flow Data'!$G$" & Row1 & ":$G$" & Row4
'Cal Flow
    ActiveChart.SeriesCollection("Calc Flow Vs Raw FLow").Values = "='Flow Data'!$J$" & Row1 & ":$J$" & Row4

End Sub

Sub MonthlyChart(ByRef startDate, ByRef endDate, ByVal startRow, ByVal endRow)
'Hao Zhang @ 2015.1.13
'this is a sub-procedure serving the changeScale()
'updates the data range and time scale for the monthly charts (both SP and TS)
Sheets(MonthName(Month(startDate), True) & " SP (Flow)").Activate
'ActiveSheet.PlotArea.Select
'level
    ActiveChart.SeriesCollection("Monitored Data").XValues = "='Flow Data'!$B$" & startRow & ":$B$" & endRow
'Flow
    ActiveChart.SeriesCollection("Monitored Data").Values = "='Flow Data'!$G$" & startRow & ":$G$" & endRow

Sheets(MonthName(Month(startDate), True) & " SP CORR (Flow)").Select
'ActiveSheet.PlotArea.Select
'level CORR
    ActiveChart.SeriesCollection("Monitored Data").XValues = "='Flow Data'!$U$" & startRow & ":$U$" & endRow
'Flow CORR
    ActiveChart.SeriesCollection("Monitored Data").Values = "='Flow Data'!$W$" & startRow & ":$W$" & endRow

Sheets(MonthName(Month(startDate), True) & " SP (Vel)").Select
'ActiveSheet.PlotArea.Select
'level
    ActiveChart.SeriesCollection("Monitored Data").XValues = "='Flow Data'!$B$" & startRow & ":$B$" & endRow
'Vel
    ActiveChart.SeriesCollection("Monitored Data").Values = "='Flow Data'!$C$" & startRow & ":$C$" & endRow

Sheets(MonthName(Month(startDate), True) & " SP CORR (Vel)").Select
'ActiveSheet.PlotArea.Select
'level CORR
    ActiveChart.SeriesCollection("Monitored Data").XValues = "='Flow Data'!$U$" & startRow & ":$U$" & endRow
'Vel CORR
    ActiveChart.SeriesCollection("Monitored Data").Values = "='Flow Data'!$V$" & startRow & ":$V$" & endRow

'''monthly TS
Sheets(MonthName(Month(startDate), True) & " TS").Select
'hyetograph
ActiveSheet.ChartObjects("Rain").Activate
'Time
    ActiveChart.SeriesCollection("Rainfall").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Rain
    ActiveChart.SeriesCollection("Rainfall").Values = "='Flow Data'!$AA$" & startRow & ":$AA$" & endRow
TSscaleChange startDate, endDate

'hydrograph
ActiveSheet.PlotArea.Select
'Dtime
    ActiveChart.SeriesCollection("Level 2").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Level2
    ActiveChart.SeriesCollection("Level 2").Values = "='Flow Data'!$D$" & startRow & ":$D$" & endRow
'Dtime
    ActiveChart.SeriesCollection("Level 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Level1
    ActiveChart.SeriesCollection("Level 1").Values = "='Flow Data'!$B$" & startRow & ":$B$" & endRow
'Dtime
    ActiveChart.SeriesCollection("Vel 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Vel1
    ActiveChart.SeriesCollection("Vel 1").Values = "='Flow Data'!$C$" & startRow & ":$C$" & endRow
'Dtime
    ActiveChart.SeriesCollection("Vel 2").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Vel2
    ActiveChart.SeriesCollection("Vel 2").Values = "='Flow Data'!$E$" & startRow & ":$E$" & endRow
'Dtime
    ActiveChart.SeriesCollection("Flow 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Flow1
    ActiveChart.SeriesCollection("Flow 1").Values = "='Flow Data'!$G$" & startRow & ":$G$" & endRow
'Dtime
    ActiveChart.SeriesCollection("Flow 2").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Flow2
    ActiveChart.SeriesCollection("Flow 2").Values = "='Flow Data'!$H$" & startRow & ":$H$" & endRow
TSscaleChange startDate, endDate

Sheets(MonthName(Month(startDate), True) & " TS CORR").Select

'hydrograph
'Dtime
    ActiveChart.SeriesCollection("Level 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Level1
    ActiveChart.SeriesCollection("Level 1").Values = "='Flow Data'!$U$" & startRow & ":$U$" & endRow
'Dtime
    ActiveChart.SeriesCollection("Flow 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Flow1
    ActiveChart.SeriesCollection("Flow 1").Values = "='Flow Data'!$W$" & startRow & ":$W$" & endRow
TSscaleChange startDate, endDate

startDate = endDate
endDate = DateAdd("m", 1, startDate)

End Sub
Sub TSscaleChange(startDate, endDate)
'Change scale
ActiveChart.Axes(xlCategory).MaximumScale = endDate
ActiveChart.Axes(xlCategory).MinimumScale = startDate
End Sub




