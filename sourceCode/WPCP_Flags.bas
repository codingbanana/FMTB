Attribute VB_Name = "WPCP_Flags"

Private Sub SWWPCP_Flag()

'******newer version embeded in SW_1min_2015_Processed.xlsm******

'Hao Zhang @ 2015.2.26
'add repeating flags for 1min WPCP data
'open the 1min raw data and run the macro

runtime = Now()
Dim mCell As range
Dim arr As Variant

'there are spaces before each heading except the first one
nams = Array("DATE_TIME", " IPS_EAST", " IPS_DELCORA", " IPS_WEST", " PLANT_DRAIN", " NETFLOW", " IPS_TOTFLOW", " IPS_CENTER")
arr = Array("Qe", "Qdel", "Qw", "Qdr", "Qt", "Ql", "Qc")

'51 = xlXMLSpreadsheet = .xlsx; 52 = xlOpenXMLWorkbookMacroEnabled = .xlsm
ActiveWorkbook.SaveAs FileFormat:=51
Set ws = ActiveWorkbook.ActiveSheet
With ws
    
    totRow = .Cells(.Rows.count, "A").End(xlUp).Row
    
    Set rng = .range("A1:H" & totRow)
    Set Key = .range("A1")
    Application.AddCustomList listArray:=nams
    .Sort.SortFields.Clear
    
'so magic!
'sort columns in customized order
    rng.Sort key1:=Key, Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=Application.CustomListCount + 1, MatchCase:=False, _
    Orientation:=xlLeftToRight, DataOption1:=xlSortNormal
    
'even more magic!!! Must add this statement after every sort, otherwise, excel will crash when saving the file!!!!!
    .Sort.SortFields.Clear
'delete customlist after sort
    Application.DeleteCustomList Application.CustomListCount
'set columns numberformat
    .range("A:A").NumberFormat = "m/d/yyyy h:mm"
    .range("B:H").NumberFormat = "General"
'add 3 columns
    .range("I1") = "Flag"
    .range("J1") = "Data Flag"
    .range("K1") = "Repeating Flag"
'add data flag
    .range("J2:J" & totRow).Formula = "=IF(IF(OR(B2<15,B2>180),""Qe"","""") & IF(OR(D2<15,D2>160),""Qw"","""") & IF(OR(C2<10,C2>120),""Qdel"","""") & IF(OR(G2<5,G2>96),""Ql"","""") & IF(F2<70,""Qt"", """") & IF(F2>600,""Qnf"", """") & IF(OR(E2<0,E2=0,E2>25),""Qd"","""") & IF(OR(H2<20,H2>300),""Qc"","""")  <>"""",IF(OR(B2<15,B2>180),""Qe"","""") & IF(OR(D2<15,D2>160),""Qw"","""") & IF(OR(C2<10,C2>120),""Qdel"","""") & IF(OR(G2<5,G2>96),""Ql"","""") & IF(F2<70,""Qt"", """") & IF(F2>600,""Qnf"", """") & IF(OR(E2<0,E2=0,E2>25),""Qdr"","""") & IF(OR(H2<20,H2>300),""Qc"",""""), ""good"")"
'add separated repeating flags, by columns
    For jCol = 2 To 8
        iRow = 2
        'get start/end row number of repeating data
        While iRow <= totRow
            startRow = iRow
            Do While .Cells(iRow + 1, jCol).Value = .Cells(iRow, jCol).Value
                iRow = iRow + 1
            Loop
            endRow = iRow
            'if repeating is longer than 5, then flag the repeating range except the first row, which is still considered as good data
            If startRow + 5 < endRow Then
                .range(.Cells(startRow + 1, jCol + 10).Address, .Cells(endRow, jCol + 10).Address).Value = arr(jCol - 2)
            End If
            iRow = iRow + 1
        DoEvents
        'show the proceedings of code
        Debug.Print "%finished = " & Round((iRow + (jCol - 2) * totRow) / (7 * totRow) * 100, 2) & "%"
        Wend
    Next
    'add consolidated repeating flags
    .range("k2:k" & .Cells(.Rows.count, "A").End(xlUp).Row).Formula = "=L2&M2&N2&O2&P2&Q2&R2"
    
    'combine data flag and repeating flag, by row
    For iRow = 2 To totRow
        If .Cells(iRow, 10).Value = "good" Then
            If .Cells(iRow, 11).Value = "" Then
                .Cells(iRow, 9).Value = "good"
            Else
                .Cells(iRow, 9).Value = .Cells(iRow, 11).Value
            End If
        Else
            If .Cells(iRow, 11).Value = "" Then
                .Cells(iRow, 9).Value = .Cells(iRow, 10).Value
            Else
                'separate repeating flag and data flag
                'Temp = .Cells(iRow, 10).Value & ","
                Temp = .Cells(iRow, 10).Value
                For jCol = 2 To 8
                    'if some repeating flags are not included in data flags, add those flags in combined flag column
                    If InStr(1, .Cells(iRow, 10).Value, .Cells(iRow, jCol + 10).Value) = 0 Then
                        Temp = Temp & .Cells(iRow, jCol + 10).Value
                    End If
                Next
                .Cells(iRow, 9).Value = Temp
                Temp = ""
           End If
        End If
    Next
End With
runtime = Now() - runtime
'http://www.globaliconnect.com/excel/index.php?option=com_content&view=article&id=121:custom-number-formats-date-a-time-formats-in-excel-a-vba-numberformat-property&catid=79&Itemid=475
'use \ after each character to display characters in literal
Debug.Print "Operation completed in " & Format(runtime, "m\m\i\n s\s\e\c") & "."
End Sub

Private Sub exportCSV_SW()
    
ActiveWorkbook.ActiveSheet.range("A:I").Copy

Set CSVwb = Workbooks.Add
CSVwb.Sheets(1).range("A1").PasteSpecial xlPasteAll

csvPath = Application.GetSaveAsFilename(fileFilter:="Excel Files (*.csv), *.csv")

If csvPath <> False Then
CSVwb.SaveAs fileName:=csvPath, FileFormat:=xlCSV
End If

End Sub

'abandoned, seems not working well
Function ExactWordInString(Text As String, Word As String) As Boolean
  ExactWordInString = " " & UCase(Text) & " " Like "*[!A-Z]" & UCase(Word) & "[!A-Z]*"
End Function
Private Sub NEWPCP_Flag()

'******newer version embeded in 5min_plantflow_2015_Processed.xlsm******

'Hao Zhang @ 2015.2.26
'add repeating flags for 1min WPCP data
'open the 1min raw data and run the macro


runtime = Timer()
Dim mCell As range
Dim arr As Variant

'there are spaces before each heading except the first one
nams = Array("DDATE", "FRANKFORD_HL", "SOMERSET_LL", "DELAWARE_LL", "DGS_PLANT_FLOW", "JCA_Radar")
arr = Array("Qh", "Ql", "Qu", "Qf", "Qa")

'51 = xlXMLSpreadsheet = .xlsx; 52 = xlOpenXMLWorkbookMacroEnabled = .xlsm
'ActiveWorkbook.SaveAs FileFormat:=51

ActiveWorkbook.ActiveSheet.Copy after:=ActiveWorkbook.ActiveSheet

Set ws = ActiveWorkbook.ActiveSheet
With ws
    ws.Name = "raw"
    totRow = .Cells(.Rows.count, "A").End(xlUp).Row
    
    Set rng = .range("A1:Q" & totRow)
    Set Key = .range("A1")
    Application.AddCustomList listArray:=nams
    .Sort.SortFields.Clear
    
'so magic!
'sort columns in customized order
    rng.Sort key1:=Key, Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=Application.CustomListCount + 1, MatchCase:=False, _
    Orientation:=xlLeftToRight, DataOption1:=xlSortNormal
    
'even more magic!!! Must add this statement after every sort, otherwise, excel will crash when saving the file!!!!!
    .Sort.SortFields.Clear
'delete customlist after sort
    Application.DeleteCustomList Application.CustomListCount
'delete extra columns
    .Columns("G:Q").Clear
'set columns numberformat
    .range("A:A").NumberFormat = "m/d/yyyy h:mm"
    .range("B:F").NumberFormat = "General"
'add duplicate flag
    .range("G1") = "Duplicate Flag"
    .range("G2") = "good"
    .range("G3:G" & totRow).Formula = "=if(A3-A2<0.003, A3-A2,""good"")"
    .range("G1:G" & totRow).AutoFilter Field:=1, Criteria1:="good"
    .Columns("A:F").Copy
End With

Set ws = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets("raw"))
With ws
    .Name = "raw+flags(no duplicate)"
    .Paste
'trim extra rows
    EOR = .Cells(.Rows.count, 1).End(xlUp).Row
    Rows(EOR & ":" & .Rows.count).Clear
'add 3 columns
    .range("G1") = "Flag"
    .range("H1") = "Data Flag"
    .range("I1") = "Repeating Flag"
'add data flag
    .range("H2:H" & EOR - 1).Formula = "=IF(IF(OR(B2=0,B2<7,B2>80),""Qh"","""") & IF(OR(C2=0,C2<14,C2>110),""Ql"","""") & IF(OR(D2=0,D2<10,D2>250),""Qu"","""") & IF(OR(E2=0,E2<50,E2>480),""Qf"","""") & IF(OR(F2<3,F2>25),""Qa"","""")<>"""", IF(OR(B2=0,B2<7,B2>80),""Qh"","""") & IF(OR(C2=0,C2<14,C2>110),""Ql"","""") & IF(OR(D2=0,D2<10,D2>250),""Qu"","""") & IF(OR(E2=0,E2<50,E2>480),""Qf"","""") & IF(OR(F2<3,F2>25),""Qa"",""""), ""good"")"
    
    '.Range("H2:H" & EOR - 1).Formula = "=IF(IF(OR(B2=0,B2<7,B2>100),""Qh"","""") & IF(OR(C2=0,C2<14,C2>140),""Ql"","""") & IF(OR(D2=0,D2<10,D2>300),""Qu"","""") & IF(OR(E2=0,E2<50,E2>500),""Qf"","""") & IF(OR(F2<3,F2>25),""Qa"","""")<>"""", IF(OR(B2=0,B2<7,B2>100),""Qh"","""") & IF(OR(C2=0,C2<14,C2>140),""Ql"","""") & IF(OR(D2=0,D2<10,D2>300),""Qu"","""") & IF(OR(E2=0,E2<50,E2>500),""Qf"","""") & IF(OR(F2<3,F2>25),""Qa"",""""), ""good"")"
'add separated repeating flags, by columns
    For jCol = 2 To 6
        iRow = 2
        'get start/end row number of repeating data
        While iRow < EOR
            startRow = iRow
            Do While .Cells(iRow + 1, jCol).Value = .Cells(iRow, jCol).Value
                iRow = iRow + 1
            Loop
            endRow = iRow
            'if repeating is longer than 5, then flag the repeating range except the first row, which is still considered as good data
            If startRow + 5 < endRow Then
                .range(.Cells(startRow + 1, jCol + 8).Address, .Cells(endRow, jCol + 8).Address).Value = arr(jCol - 2)
            End If
            iRow = iRow + 1
        DoEvents
        'show the proceedings of code
        Debug.Print "%finished = " & Round((iRow + (jCol - 2) * totRow) / (5 * totRow) * 100, 2) & "%"
        Wend
    Next
    
    'since UDLL is calculated from Total Flow, when TotFlow is flaged, so should UDLL.
    For iRow = 2 To EOR
        If .Cells(iRow, "M") = "Qf" And .Cells(iRow, "L") <> "Qu" Then
            .Cells(iRow, "L") = "Qu"
        End If
    Next
    
    'add consolidated repeating flags
    .range("I2:I" & .Cells(.Rows.count, "A").End(xlUp).Row).Formula = "=J2&K2&L2&M2&N2"
    
    'combine data flag and repeating flag, by row
    For iRow = 2 To EOR - 1
        If .Cells(iRow, 8).Value = "good" Then
            If .Cells(iRow, 9).Value = "" Then
                .Cells(iRow, 7).Value = "good"
            Else
                .Cells(iRow, 7).Value = .Cells(iRow, 9).Value
            End If
        Else
            If .Cells(iRow, 9).Value = "" Then
                .Cells(iRow, 7).Value = .Cells(iRow, 8).Value
            Else
                'separate repeating flag and data flag
                'Temp = .Cells(iRow, 10).Value & ","
                Temp = .Cells(iRow, 8).Value
                For jCol = 2 To 6
                    'if some repeating flags are not included in data flags, add those flags in combined flag column
                    If InStr(1, .Cells(iRow, 8).Value, .Cells(iRow, jCol + 8).Value) = 0 Then
                        Temp = Temp & .Cells(iRow, jCol + 8).Value
                    End If
                Next
                .Cells(iRow, 7).Value = Temp
                Temp = ""
           End If
        End If
    Next
End With
runtime = Round(Timer() - runtime, 1)
'http://www.globaliconnect.com/excel/index.php?option=com_content&view=article&id=121:custom-number-formats-date-a-time-formats-in-excel-a-vba-numberformat-property&catid=79&Itemid=475
'use \ after each character to display characters in literal
Debug.Print "Operation completed in " & runtime & " sec."
End Sub

Private Sub exportCSV_NE()
    
ActiveWorkbook.Sheets("raw+flag(no duplicate)").range("A:G").Copy

Set CSVwb = Workbooks.Add
CSVwb.Sheets(1).Paste
'CSVwb.Sheets(1).Range("A1").PasteSpecial xlPasteAll

csvPath = Application.GetSaveAsFilename(fileFilter:="Excel Files (*.csv), *.csv")

If csvPath <> False Then
CSVwb.SaveAs fileName:=csvPath, FileFormat:=xlCSV
End If

End Sub

