Attribute VB_Name = "QAQC"
Sub QC_multi_file()
'Hao Zhang @ 2014/11/18
'this macro executes following steps:
'1.get the site name from QA logbook.xlsm
'2.open the 3 source files from server
'3.copy the template from server to destination folder
'4.copy data from source files and paste to designated cells in template
'5.leave a note on QA Logbook.xlsm for the completed site
'6.move to the next site

'Part of the code is inspired from Eddie's Macro Pipe_Profile_Checker_Macro()

Static i As Integer
'Stop the screen from flashing
    Application.ScreenUpdating = False

'initialize variables
    Dim path As String
    Dim ManholeID As String
    Dim source_LVF As String
    Dim source_Redun As String
    Dim source_FP As String
    Dim source_LVF_Var As String
    Dim source_Redun_Var As String
    Dim source_FP_Var As String
    Dim target As String
    Dim file_LVF As String
    Dim file_Redun As String
    Dim file_FP As String
    Dim file_target As String
    Dim list_Len As Integer
    Dim lastRow As Integer
    Dim rawFolder As String
    Dim rawDate As String
    Dim startTime As Date
    Dim endTime As Date
    
    Dim fso
    Dim SFolder As String
    rawFolder1 = "PE 2014-10-15"
    rawDate1 = "(2014-11-1 to 2014-11-15)"
    rawFolder2 = "PE 2014-11-30"
    rawDate2 = "(2014-11-16 to 2014-12-01)"
    
'check if the index page is already open, otherwise, open it
    Dim Ret

    Ret = IsWorkBookOpen("C:\Users\hao.zhang\Desktop\QA logbook.xlsm")

    If Ret = True Then
        Workbooks("QA logbook.xlsm").Activate
    Else
        Workbooks.Open fileName:="C:\Users\hao.zhang\Desktop\QA logbook.xlsm", UpdateLinks:=0
    End If
'find the end of list
    list_Len = Sheets("log").UsedRange.Rows.count
'Manually loop through the list from B4 on QA logbook.xlsm/log
'set the initial value of i to 4
    If i = 0 Then
        i = 4
    End If
'get manholeID and path for the site
        ManholeID = Sheets("log").range("B" & i).Value
        path = Sheets("log").range("D" & i).Value
'there are two format of raw data folder, and two format of primary excel filenames
        path_ext = path & "\Raw Data\" & rawFolder & "\"
        path_ext_Var = path & "\RawData\" & rawFolder & "\"
        file_LVF = ManholeID & " - Excel " & rawDate & ".csv"
        file_LVF_Var = ManholeID & " - Excel With Watertemp " & rawDate & ".csv"
        file_Redun = ManholeID & " - Redundant Excel " & rawDate & ".csv"
        file_FP = ManholeID & " Electronic Fieldbook.xls"
'use a control structure to decide the file path/name
Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path_ext) Then
        source_Redun = path_ext & file_Redun    'get file path of redundant level data
        source_FP = path_ext & file_FP          'get file path of field points
            If Dir(path_ext & file_LVF) Then
                source_LVF = path_ext & file_LVF 'get file path of level, velocity, flow data
            Else
                source_LVF = path_ext & file_LVF_Var
            End If
    Else
        source_Redun_Var = path_ext_Var & file_Redun
        source_FP_Var = path_ext_Var & file_FP
        If Dir(path_ext_Var & file_LVF) Then
        source_LVF = path_ext_Var & file_LVF
        Else
        source_LVF = path_ext_Var & file_LVF_Var
        End If
    End If

'get file path of previous quarterly summary
    source_Q3 = path & "\QAQC\" & ManholeID & " (Q3-14).xlsm"
'get QAQC file path
    file_target = ManholeID & " (Q4-14).xlsm"
    target = path & "\QAQC\" & file_target

'0. open QA file,create a new one if not exist
If Dir(target) <> "" Then
Workbooks.Open fileName:=target, UpdateLinks:=0
Else
FileCopy "\\pwdhqr\oows\Modeling\Data\Temporary Monitors\Flow Monitoring\Supplementary Documents\QAQC_SSOAP_R-value_Templates\QAQC\2014\Template (Q4-14).xlsm", target
End If

'1. open source_LVF file
Workbooks.Open fileName:=source_LVF, UpdateLinks:=0

'copy level, velocity data from source_LVF file

''''''''''''''''''''''''''''''''debug stopped here''''''''''''''''''''''''''''''''''

range("B2:C1537").Select
 Application.CutCopyMode = False
Selection.Copy
'paste level, velocity data to QA file - Flow Data sheet
Workbooks(file_target).Activate
Sheets("Flow data").Select
range("B1455:C2986").Select
ActiveSheet.Paste

'2. copy flow data from source_LVF file
Workbooks(file_LVF).Activate
range("D2:D1537").Select
 Application.CutCopyMode = False
Selection.Copy
'open QA file
Workbooks(file_target).Activate
'paste flow data to QA file - Flow Data sheet
Sheets("Flow data").Activate
range("G1455:G2986").Select
ActiveSheet.Paste

'3. copy Temperature data from source_LVF file
Workbooks(file_LVF).Activate
range("E2:E1537").Select
 Application.CutCopyMode = False
Selection.Copy
'open QA file
Workbooks(file_target).Activate
'paste Temperature data to QA file - Flow Data sheet
Sheets("Flow data").Activate
range("F1455:F2986").Select
ActiveSheet.Paste

'4. open redundant level file
Workbooks.Open fileName:=source_Redun, UpdateLinks:=0
'copy redundant level data from source_LVF file
range("B2:B1537").Select
 Application.CutCopyMode = False
Selection.Copy
'open QA file
Workbooks(file_target).Activate
'paste redundant level data to QA file - Flow Data sheet
Sheets("Flow data").Activate
range("D1455:D2986").Select
ActiveSheet.Paste

'5. open electronic fieldbook file
Workbooks.Open fileName:=source_FP, UpdateLinks:=0
'find out the last row of data
range("B2:D2").Select
Selection.End(xlDown).Select
lastRow = ActiveCell.Row

'copy field point: time from source_FP file
range("B4:D" & lastRow).Select
 Application.CutCopyMode = False
Selection.Copy
'open QA file
Workbooks(file_target).Activate
'paste time to QA file - Site Info
Sheets("Site Info").Activate
range("C23").Select
ActiveSheet.Paste

'6. copy field point:flow data from source_FP file
Workbooks(file_FP).Activate
range("E4:E" & lastRow).Select
 Application.CutCopyMode = False
Selection.Copy
'open QA file
Workbooks(file_target).Activate
'paste flow data to QA file - Site Info
Sheets("Site Info").Activate
range("H23").Select
ActiveSheet.Paste

'7. copy field point:depth data from source_FP file
Workbooks(file_FP).Activate
range("H4:K" & lastRow).Select
 Application.CutCopyMode = False
Selection.Copy
'open QA file
Workbooks(file_target).Activate
'paste flow data to QA file - Site Info
Sheets("Site Info").Activate
range("K23").Select
ActiveSheet.Paste

'8. copy field point:comment data from source_FP file
Workbooks(file_FP).Activate
range("N4:R" & lastRow).Select
 Application.CutCopyMode = False
Selection.Copy
'open QA file
Workbooks(file_target).Activate
'paste flow data to QA file - Site Info
Sheets("Site Info").Activate
range("Q23").Select
ActiveSheet.Paste
Cells(2, 2).Value = Now()
'9. add a comment on the QA logbook.xlsm
Workbooks("QA logbook.xlsx").Activate
Sheets("log2").Activate
Cells(i, 3).Value = "QA sheets updated to 10/31/2014"

'Next i

End Sub
Function IsWorkBookOpen(fileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open fileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function
Sub QC_Q4_2014_convert()
'Hao Zhang @ 2014/11/14
'This macro automates the transformation from old template to new template
'
'
Dim FieldPointLastRow As Integer
 Const sourceFile = "BC-B0755 (Q4-14).xlsm"
 Const targetFile = "Template (Q4-14).xlsm"
'copy site info
    Windows(sourceFile).Activate
    Sheets("Site Info").Select
    range("B3:C20").Select
    Selection.Copy
'paste site info
    Windows(targetFile).Activate
    Sheets("Site Info").Select
    range("B3").Select
    ActiveSheet.Paste
'add last modified date and initial
    range("B2").Select
    ActiveCell.Value = Now()
    range("C2").Select
    ActiveCell.Value = "HZ"
'copy field points (part 1)
    Windows(sourceFile).Activate
    range("C23:H23").Select
    range(Selection, Selection.End(xlDown)).Select
    FieldPointLastRow = Selection.End(xlDown).Row
    Selection.Copy
'paste field points (part 1)
    Windows(targetFile).Activate
    range("C23").Select
    ActiveSheet.Paste
'copy field points (part 2)
    Windows(sourceFile).Activate
    range("K23:N" & FieldPointLastRow).Select
    Application.CutCopyMode = False
    Selection.Copy
'paste field points (part 2)
    Windows(targetFile).Activate
    range("K23").Select
    ActiveSheet.Paste
'copy field points (part 3)
    Windows(sourceFile).Activate
    range("Q23:U" & FieldPointLastRow).Select
    Application.CutCopyMode = False
    Selection.Copy
'paste field points (part 3)
    Windows(targetFile).Activate
    range("Q23").Select
    ActiveSheet.Paste
'copy area vs depth data
    Windows(sourceFile).Activate
    Sheets("Area vs. depth table").Select
    range("A3").Select
    range(Selection, Selection.End(xlToRight)).Select
    range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
 'paste area vs depth data
    Windows(targetFile).Activate
    Sheets("Area vs. depth table").Select
    range("A3").Select
    ActiveSheet.Paste
 'copy flow data
    Windows(sourceFile).Activate
    Sheets("Flow Data").Select
    range("B15:H15").Select
    range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
 'paste flow data
    Windows(targetFile).Activate
    Sheets("Flow Data").Select
    range("B15").Select
    ActiveSheet.Paste
  'clear empty formula cells in flow data so figures can be plotted correctly
    range("W1455:Z1455").Select
    range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
  'copy rainfall data
    Sheets("Rainfall Data").Select
    Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy h:mm"
    range("B2").Select
    range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
 'paste rainfall data into flow data worksheet
    Sheets("Flow Data").Select
    range("AC15").Select
    ActiveSheet.Paste
End Sub

Sub QC_Year_shift()
'Hao Zhang @ 2014/11/14
'Corrects rainfall time range on ALL TS and ALL TS CORR tab in Q4-2014 QA sheets.

Sheets("ALL TS").Select
ActiveSheet.ChartObjects("Chart 51").Activate
ActiveChart.Axes(xlCategory).MaximumScale = 42005
ActiveChart.Axes(xlCategory).MinimumScale = 41913
'change rainfall plot on ALL TS CORR tab
Sheets("ALL TS CORR").Select
ActiveSheet.ChartObjects("Chart 51").Activate
ActiveChart.Axes(xlCategory).MaximumScale = 42005
ActiveChart.Axes(xlCategory).MinimumScale = 41913
End Sub

Sub QC_Year_shift_Temp()
'Hao Zhang @ 2014/11/14
'Correct rainfall time range on Temp TS CORR and CORR Temp TS CORR tab issue existed in Q4-2014 QA sheets.
'this is a variation of the QC_Year_shift() that deals with temperature plots

Sheets("Temp TS CORR").Select
ActiveSheet.ChartObjects("Chart 51").Activate
ActiveChart.Axes(xlCategory).MaximumScale = 42005   '42005 = '1/1/2015'
ActiveChart.Axes(xlCategory).MinimumScale = 41913   '41913 = '10/1/2014'
'change rainfall plot on ALL TS CORR tab
Sheets("CORR Temp TS CORR ").Select
ActiveSheet.ChartObjects("Chart 51").Activate
ActiveChart.Axes(xlCategory).MaximumScale = 42005
ActiveChart.Axes(xlCategory).MinimumScale = 41913
End Sub

Sub yr_sh1ft()
'Hao Zhang @ 2014.11.20
'Correct the time range of Oct, Nov, Dec TS and TS CORR chart-hyetograph
'go to the monthly precipitation chart first
ActiveSheet.ChartObjects("Chart 98").Activate
ActiveChart.Axes(xlCategory).MaximumScale = 41944   '41944 = '11/1/2014'
ActiveChart.Axes(xlCategory).MinimumScale = 41913   '41913 = '10/1/2014'
End Sub
Sub Month_TS_chart()
'Hao Zhang @2015.1.8
'set the time range of Oct Nov, Dec TS and TS CORR chart


'let user select quarter, then set the start time
Select Case QAQC_form.QuarterCbox.Value
Case Is = "Q1 (Jan-Mar)"
    startMonth = 1
    month1 = "Jan"
    Month2 = "Feb"
    Month3 = "Mar"
Case Is = "Q2 (Apr-Jun)"
    startMonth = 4
    month1 = "Apr"
    Month2 = "May"
    Month3 = "Jun"
Case Is = "Q3 (Jul-Sept)"
    startMonth = 7
    month1 = "Jul"
    Month2 = "Aug"
    Month3 = "Sept"
Case Is = "Q4 (Oct-Dec)"
    startMonth = 10
    month1 = "Oct"
    Month2 = "Nov"
    Month3 = "Dec"
End Select

mYear = QAQC_form.yearCbox.Value

startDate = DateSerial(mYear, startMonth, 1)
endDate = DateAdd("m", 1, startDate)

'first month TS
Sheets(MonthName(Month(startDate), True) & " TS").Select
'hyetograph
ActiveSheet.ChartObjects("Chart 98").Activate
ActiveChart.Axes(xlCategory).MaximumScale = endDate
ActiveChart.Axes(xlCategory).MinimumScale = startDate
startDate = endDate
endDate = DateAdd("m", 1, startDate)

'second month TS
Sheets(MonthName(Month(startDate), True) & " TS").Select
'hyetograph
ActiveSheet.ChartObjects("Chart 82").Activate
ActiveChart.Axes(xlCategory).MaximumScale = endDate
ActiveChart.Axes(xlCategory).MinimumScale = startDate
startDate = endDate
endDate = DateAdd("m", 1, startDate)

'third month TS
Sheets(MonthName(Month(startDate), True) & " TS").Select
'hyetograph
ActiveSheet.ChartObjects("Chart 83").Activate
ActiveChart.Axes(xlCategory).MaximumScale = endDate
ActiveChart.Axes(xlCategory).MinimumScale = startDate

'all TS
Sheets("ALL TS").Select
'hyetograph
ActiveSheet.ChartObjects("Chart 51").Activate
ActiveChart.Axes(xlCategory).MaximumScale = endDate
ActiveChart.Axes(xlCategory).MinimumScale = DateAdd("m", -3, endDate)
'all TS CORR
Sheets("ALL TS CORR").Select
'hyetograph
ActiveSheet.ChartObjects("Chart 51").Activate
ActiveChart.Axes(xlCategory).MaximumScale = endDate
ActiveChart.Axes(xlCategory).MinimumScale = DateAdd("m", -3, endDate)

End Sub
Sub extendFPrange()
'Hao Zhang @2015.1.12
'fix the bug in QA sheets that FP points are only plotted up to row 147
'new end of row = 300
'let user select quarter, then set the start time
Select Case QAQC_form.QuarterCbox.Value
Case Is = "Q1 (Jan-Mar)"
    startMonth = 1
    month1 = "Jan"
    Month2 = "Feb"
    Month3 = "Mar"
Case Is = "Q2 (Apr-Jun)"
    startMonth = 4
    month1 = "Apr"
    Month2 = "May"
    Month3 = "Jun"
Case Is = "Q3 (Jul-Sept)"
    startMonth = 7
    month1 = "Jul"
    Month2 = "Aug"
    Month3 = "Sept"
Case Is = "Q4 (Oct-Dec)"
    startMonth = 10
    month1 = "Oct"
    Month2 = "Nov"
    Month3 = "Dec"
End Select

mYear = QAQC_form.yearCbox.Value

startDate = DateSerial(mYear, startMonth, 1)
endDate = DateAdd("m", 1, startDate)
'the macro must be running from a worksheet rather than a chart,
'otherwise, it will throw out a runtime error 1004
ActiveWorkbook.Sheets(1).Select

'first month SP (Flow)
For i = 1 To 3
Sheets(MonthName(Month(startDate), True) & " SP (Flow)").Activate
Call extendFP_SP_Flow
Sheets(MonthName(Month(startDate), True) & " SP CORR (Flow)").Select
Call extendFP_SP_Flow
Sheets(MonthName(Month(startDate), True) & " SP (Vel)").Select
''only this chart is treated differently because it uses different series name
ActiveSheet.PlotArea.Select
'FP level
    ActiveChart.SeriesCollection("VP").XValues = "='Site Info'!$H$23:$H$300"
'FP velocity
    ActiveChart.SeriesCollection("VP").Values = "='Site Info'!$K$23:$K$300"
Sheets(MonthName(Month(startDate), True) & " SP CORR (Vel)").Select
Call extendFP_SP_Vel
Sheets(MonthName(Month(startDate), True) & " TS").Select
Call extendFP_TS
startDate = endDate
endDate = DateAdd("m", 1, startDate)
Next i

'all TS
Sheets("ALL TS").Select
Call extendFP_TS
'SP Velocity Vs Level 1&2
Sheets("SP Velocity Vs Level 1&2").Select
Call extendFP_SP_Vel
'SP Flow Vs Level 1&2
Sheets("SP Flow Vs Level 1&2").Select
Call extendFP_SP_Flow

End Sub
Sub extendFP_SP_Flow()
ActiveSheet.PlotArea.Select
'FP level
    ActiveChart.SeriesCollection("FP").XValues = "='Site Info'!$H$23:$H$300"
'FP Flow
    ActiveChart.SeriesCollection("FP").Values = "='Site Info'!$J$23:$J$300"
End Sub
Sub extendFP_SP_Vel()
ActiveSheet.PlotArea.Select
'FP level
    ActiveChart.SeriesCollection("FP").XValues = "='Site Info'!$H$23:$H$300"
'FP velocity
    ActiveChart.SeriesCollection("FP").Values = "='Site Info'!$K$23:$K$300"
End Sub
Sub extendFP_TS()
ActiveSheet.PlotArea.Select
'FP time
    ActiveChart.SeriesCollection("Field Level").XValues = "='Site Info'!$B$23:$B$300"
'FP level
    ActiveChart.SeriesCollection("Field Level").Values = "='Site Info'!$H$23:$H$300"
'FP time
    ActiveChart.SeriesCollection("Field Velocity").XValues = "='Site Info'!$B$23:$B$300"
'FP velocity
    ActiveChart.SeriesCollection("Field Velocity").Values = "='Site Info'!$K$23:$K$300"
'FP time
    ActiveChart.SeriesCollection("Field Flow").XValues = "='Site Info'!$B$23:$B$300"
'FP flow
    ActiveChart.SeriesCollection("Field Flow").Values = "='Site Info'!$J$23:$J$300"
'FP time
    ActiveChart.SeriesCollection("Silt").XValues = "='Site Info'!$B$23:$B$300"
'FP silt
    ActiveChart.SeriesCollection("Silt").Values = "='Site Info'!$R$23:$R$300"


End Sub


Sub temp_chart()
'Hao Zhang @2015.1.8
'Adjust chart CORR Temp TS and CORR Temp TS CORR data range
'The reason of doing this is because when coping charts from another workbook,
'the data series are still linked to the source file.
'By changing it to the local ranges, the plot can correctly represents the data

'Before runing the macro, copy both charts from another QA sheet.
'after runing the macro, do [Data tab]-> Edit Links-> Break Links-> yes

Sheets("Temp TS CORR").Select
'change hyteograph
    ActiveSheet.ChartObjects("Chart 51").Activate 'Chart 51= hyetograph
'    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$A$15:$A$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$AC$15:$AC$8846"
'    ActiveChart.Axes(xlCategory).MaximumScale = QAQC_form.EndTimeTextBox.value
'    ActiveChart.Axes(xlCategory).MinimumScale = QAQC_form.startTimeTextBox.value
'change hydrograph
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$A$15:$A$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$W$15:$W$8846"
    ActiveChart.SeriesCollection(2).XValues = "='Flow Data'!$A$15:$A$8846"
    ActiveChart.SeriesCollection(2).Values = "='Flow Data'!$F$15:$F$8846"
    
Sheets("CORR Temp TS CORR").Select
'change hyteograph
    ActiveSheet.ChartObjects("Chart 51").Activate
'    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$A$15:$A$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$AC$15:$AC$8846"
'    ActiveChart.Axes(xlCategory).MaximumScale = QAQC_form.EndTimeTextBox.Value
'    ActiveChart.Axes(xlCategory).MinimumScale = QAQC_form.startTimeTextBox.Value

'change hydrograph
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$A$15:$A$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$W$15:$W$8846"
    ActiveChart.SeriesCollection(2).XValues = "='Flow Data'!$A$15:$A$8846"
    ActiveChart.SeriesCollection(2).Values = "='Flow Data'!$Z$15:$Z$8846"
End Sub

Sub Rainlink()
'Hao Zhang @2015.1.9
'pull rainfall data from PWD2010.mdb without generate additional table
'the code was adapted from [SSOAP].rain_query()

    Dim RG As Integer
    Dim startTime As Date, endTime As Date
    Dim ws As Worksheet
    Dim Cn As ADODB.Connection, rs As ADODB.Recordset
    Dim MyConn, sSQL As String
    Dim iCol, vArr, EOR

    RG = QAQC_form.RGTextBox.Value
    
    startTime = QAQC_form.startTimeTextBox.Value
    endTime = QAQC_form.EndTimeTextBox.Value
        
'set the worksheet if exist, otherwise, create one then set it
On Error Resume Next
If ActiveWorkbook.Sheets("Rainfall Data") Is Nothing Then
    If ActiveWorkbook.Sheets("Rainfall") Is Nothing Then
        Set ws = ActiveWorkbook.Sheets.Add(after:=Sheets(ActiveWorkbook.Sheets.count))
        ws.Name = "Rainfall Data"
    Else
        Set ws = ActiveWorkbook.Sheets("Rainfall")
        ws.Name = "Rainfall Data"
    End If
Else
    Set ws = ActiveWorkbook.Sheets("Rainfall Data")
End If

     'Set source
    MyConn = "C:\Rainfall\PWDRAIN2010\PWDRAIN2010.mdb"
     'Create query
    sSQL = "SELECT Daytime, finalRG" & RG & " FROM [FinalAll(2014)] WHERE (((Daytime) >= #" & startTime & "# And (Daytime) < #" & endTime & "#));"
    
     'Create RecordSet
    Set Cn = New ADODB.Connection
    With Cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"  'ACE is a newer and better oleDB driver than JET
'       .Provider = "Microsoft.Jet.OLEDB.4.0"
        .CursorLocation = adUseClient
        .Open MyConn
        Set rs = .Execute(sSQL)
    End With
    
    'Clear previous results
    ws.range("A:B").ClearContents
    'get the titles, this is the fancy way, which can be used if varied columns are involved
    For icols = 0 To rs.Fields.count - 1
        ws.Cells(1, icols + 1).Value = rs.Fields(icols).Name
    Next

    'set title font to be bold
        ws.range(ws.Cells(1, 1), ws.Cells(1, rs.Fields.count)).Font.Bold = True
    
    'Write RecordSet to results area
        ws.range("A2").CopyFromRecordset rs
    
    'release the object
    rs.Close
    Cn.Close
    Set Cn = Nothing

Dim fd As Worksheet
Set fd = ActiveWorkbook.Sheets("Flow Data")
'find the last row of the quarter in Flow Data
EOR = fd.Cells(range("A:A").Rows.count, "A").End(xlUp).Row
'find the column for rainfall in Flow Data
rainCol = fd.range("A13:AZ13").Find("Rain Fall Data").Column
'convert the column number back to letter
vArr = Split(Cells(1, rainCol).Address(True, False), "$")
'link Flow data!rainfall data (column AC) to Rainfall!rainfall(column B)

fd.Activate
fd.Cells(15, rainCol).Select
If Selection.Formula <> "='Rainfall Data'!B2" Then
    Selection.Value = "='Rainfall Data'!B2"
    Selection.AutoFill Destination:=range(vArr(0) & "15", vArr(0) & EOR)
End If

End Sub

Sub trim_tail()
'Hao Zhang @2015.1.9
'autofill formula for the Corrected Column, and delete extra formula so charts can be shown properly

'find the end of row (EOR)
Dim ws As Worksheet
Set ws = ActiveWorkbook.Sheets("Flow Data")
ws.Activate
EOR = ws.Cells(range("B:B").Rows.count, 2).End(xlUp).Row

'autofill to the EOR
'****this is a terrible mistake, it erases deleted data!!!
'ws.Range("W15", "Z15").AutoFill Destination:=ws.Range("W15", "Z" & EOR), Type:=xlFillDefault
'clearcontent after the EOR
ws.range("W" & EOR + 1, "Z" & ws.Rows.count).ClearContents

End Sub


Sub NovDecChart()
'
' Hao Zhang @ 2015.1.20
'target file: Q4-14.xlsm
'issue: Dec charts are starting at 12/2/2014, while Nov charts are ending at 12/2/2014
'action: adjust Nov and Dec charts to proper time range


    Sheets("Nov SP (Flow)").Select
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$B$2991:$B$5870"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$G$2991:$G$5870"
    Sheets("Nov SP CORR (Flow)").Select
    ActiveChart.ChartArea.Select
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$Y$2991:$Y$5870"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$W$2991:$W$5870"
    Sheets("Nov SP (Vel)").Select
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$B$2991:$B$5870"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$C$2991:$C$5870"
    Sheets("Nov SP CORR (Vel)").Select
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$Y$2991:$Y$5870"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$X$2991:$X$5870"
    Sheets("Dec SP (Flow)").Select
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$B$5871:$B$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$G$5871:$G$8846"
    Sheets("Dec SP CORR (Flow)").Select
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$Y$5871:$Y$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$W$5871:$W$8846"
    Sheets("Dec SP (Vel)").Select
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$B$5871:$B$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$C$5871:$C$8846"
    Sheets("Dec SP CORR (Vel)").Select
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$Y$5871:$Y$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$X$5871:$X$8846"
    Sheets("Dec TS").Select
    ActiveSheet.ChartObjects("Chart 83").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$A$5871:$A$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$AC$5871:$AC$8846"
    ActiveSheet.ChartObjects("Chart 83").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection("Level 2").XValues = "='Flow Data'!$A$5871:$A$8846"
    ActiveChart.SeriesCollection("Level 2").Values = "='Flow Data'!$D$5871:$D$8846"
    ActiveChart.SeriesCollection("Level 1").XValues = "='Flow Data'!$A$5871:$A$8846"
    ActiveChart.SeriesCollection("Level 1").Values = "='Flow Data'!$B$5871:$B$8846"
    ActiveChart.SeriesCollection("Vel 1").XValues = "='Flow Data'!$A$5871:$A$8846"
    ActiveChart.SeriesCollection("Vel 1").Values = "='Flow Data'!$C$5871:$C$8846"
    ActiveChart.SeriesCollection("Vel 2").XValues = "='Flow Data'!$A$5871:$A$8846"
    ActiveChart.SeriesCollection("Vel 2").Values = "='Flow Data'!$E$5871:$E$8846"
    ActiveChart.SeriesCollection("Flow 1").XValues = "='Flow Data'!$A$5871:$A$8846"
    ActiveChart.SeriesCollection("Flow 1").Values = "='Flow Data'!$G$5871:$G$8846"
    ActiveChart.SeriesCollection("Flow 2").XValues = "='Flow Data'!$A$5871:$A$8846"
    ActiveChart.SeriesCollection("Flow 2").Values = "='Flow Data'!$H$5871:$H$8846"
    
    Sheets("Dec TS CORR").Select
    ActiveChart.ChartArea.Select
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(1).XValues = "='Flow Data'!$A$5871:$A$8846"
    ActiveChart.SeriesCollection(1).Values = "='Flow Data'!$Y$5871:$Y$8846"
    ActiveChart.SeriesCollection(2).XValues = "='Flow Data'!$A$5871:$A$8846"
    ActiveChart.SeriesCollection(2).Values = "='Flow Data'!$W$5871:$W$8846"
End Sub

