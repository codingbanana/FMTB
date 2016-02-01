Attribute VB_Name = "BigPicture"
Dim fso As New FileSystemObject

Sub BigPicture()
'Hao Zhang @ 2014.11.21
'Hao Zhang revised@2014.12.31
'prepare datasheets for bigPicture analysis following steps below:
'1. copy most recent 'Combined_field_points.csv', 'CombinedQAQC.csv' to local harddrive in new names incl. date and initial
'2. copy 'input.csv' and 'Run_BigPicture.bat' to local drive
'3. move the two existing files in step 1 to bk folder on server
'4. open the most recent quarterly QA sheet, append dtime, level, flow, velocity, corrected.flow, corrected.level to the CSVs
'5. change the file name and path in input.csv
'6. (manual) run Run_BigPicture.bat
'7. (future) prompt a msgbox to check results
'8. move generated big_picture back to server, overwrite existing version
'9. start working on the next site
'10. (future) make the code robust for other quarter/year
'keyword: create folder (mkdir), copy & paste file, partial match file name
'(dir), move file, For-loop


'PART 0 : Initialization
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Goto ActiveCell, True

Dim sites As String
Dim sFound1 As String
Dim sFound2 As String
Dim source1 As String
Dim source2 As String
Dim target1 As String
Dim target2 As String
Dim target3 As String
Dim target4 As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim Ret As Boolean
Dim wb0 As Workbook
Dim wb1 As Workbook
Dim wb2 As Workbook
Dim wb3 As Workbook
Dim ws0 As Worksheet
Dim ws1 As Worksheet
Dim ws1b As Worksheet
Dim ws2 As Worksheet
Dim ws3 As Worksheet

'open VBA-MAcro.xlsm first

'check if QA Logbook has been opened or not
''' this is a better way to inqury if a file/sheet is open
Set wb0 = Workbooks("QA Logbook.xlsm")
If wb0 Is Nothing Then
Set wb0 = Workbooks.Open("C:\Users\hao.zhang\Desktop\QA Logbook.xlsm")
End If

'CreateTab() was here
'to simplify the process, manually create a new tab from previous BigP tab
'the macro will use whatever is activated as the reference worksheet
Set ws0 = wb0.ActiveSheet

'start the loop that go through each site from the first row that don't have notes till the first blank row
For i = ws0.Cells(Rows.count, 6).End(xlUp).Row + 1 To ws0.range("c1").End(xlDown).Row
    sites = ws0.Cells(i, 3).Value
    localpath = "C:\Users\hao.zhang\Desktop\BigPicture\" & sites & "\"
'create a sub-folder inside BigPicture with siteID as the folder name
'BigPicture folder must be existed already
'#flow-control needed to avoid errors when folders already existed
   On Error Resume Next
        MkDir (localpath)

'adjust windows for visual check
ws0.Activate
With ActiveWindow
    .Width = 800
    .Height = 300
    .Top = 0
    .Left = 0
    .ScrollRow = i - 10
    .ScrollColumn = 1
End With

'find the first file that partially matches the source file name
'only the filename with extension is assigned to sFound1 and sFound2
' * is used as a wildcard
'set the full path of source files
sFound1 = Dir(ws0.Cells(i, 4).Value & "QAQC\BigPicture\" & sites & "_Combined_field_points*.csv")
source1 = Cells(i, 4).Value & "QAQC\BigPicture\" & sFound1
source3 = Cells(i, 4).Value & "QAQC\BigPicture\input.csv"
If sFound1 = "" Then
sFound1 = Dir(ws0.Cells(i, 4).Value & "QAQC\Big Picture\" & sites & "_Combined_field_points*.csv")
source1 = Cells(i, 4).Value & "QAQC\Big Picture\" & sFound1
source3 = Cells(i, 4).Value & "QAQC\Big Picture\input.csv"
End If

sFound2 = Dir(ws0.Cells(i, 4).Value & "QAQC\BigPicture\" & sites & "_CombinedQAQC*.csv")
source2 = Cells(i, 4).Value & "QAQC\BigPicture\" & sFound2
If sFound2 = "" Then
sFound2 = Dir(ws0.Cells(i, 4).Value & "QAQC\Big Picture\" & sites & "_CombinedQAQC*.csv")
source2 = Cells(i, 4).Value & "QAQC\Big Picture\" & sFound2
End If

'set the full path of target files
target1 = localpath & sites & "_Combined_field_points_" & Format(Now(), "yymmdd") & "_HZ.csv"
target2 = localpath & sites & "_CombinedQAQC_" & Format(Now(), "yymmdd") & "_HZ.csv"
target3 = localpath & "input.csv"
target4 = localpath & "Run_BigPicture.bat"

'used the updated bat file from local harddrive instead (changed R version, server name)
source4 = "C:\Users\hao.zhang\Desktop\BigPicture\Run_BigPicture.bat"
'set the full path of bk files (for storing old files)
bk1 = Cells(i, 4).Value & "QAQC\BigPicture\bk\" & sFound1
bk2 = Cells(i, 4).Value & "QAQC\BigPicture\bk\" & sFound2
'set the full path of QA sheets
''####consider changing to partial match for better robustness
QA_sheet = ws0.range("E" & i).Hyperlinks.Item(1).Address
QA_sheet_Q3 = ws0.Cells(i, 4).Value & "QAQC\" & sites & " (Q3-14).xlsx"


'PART I: File Handling

'copy Combined_FP from server to local hard drive
On Error Resume Next    'if file exists, move on
If sFound1 <> "" Then
    FileCopy source1, target1
'copy existing Combined_FP file to bk folder
'    FileCopy source1, bk1
'move existing Combined_FP file to bk folder (temporaly suspended)
    Name source1 As bk1
Else
    MsgBox ("can't find FP file!")
    Exit Sub
End If

'copy Combined_QAQC from server to local hard drive
If sFound2 <> "" Then
    FileCopy source2, target2
'copy existing Combined_FP file to bk folder
'    FileCopy source2, bk2
'move existing Combined_QAQC to bk folder (temporaly suspended)
    Name source2 As bk2
Else
    MsgBox ("can't find the QAQC file!")
    Exit Sub
End If

'copy input from server to local hard drive
FileCopy source3, target3

'copy Run_bigPicture from BigPicutre folder to local hard drive
FileCopy source4, target4

'open QA_sheet
DoEvents
Workbooks.Open fileName:=QA_sheet
Set wb1 = Workbooks(Dir(QA_sheet))
Set ws1 = wb1.Sheets("site info")
Set ws1b = wb1.Sheets("Flow data")

'adjust windows for visual check
ws1.Activate
With ActiveWindow
    .Width = 400
    .Height = 800
    .Top = 300
    .Left = 400
    .ScrollColumn = 2
End With

'PART II : add data to FP file (ws1->ws2)

'open FP file on local drive
Workbooks.Open fileName:=target1
Set wb2 = Workbooks(Dir(target1))
Set ws2 = wb2.Sheets(1)

'find the columns number of dTime, level, flow, and velocity
With ws1.range("A22:V22")
y0 = .Find("Date Time").Column
y1 = .Find(" Field Level (inches)").Column
y2 = .Find("Field Flow (mgd)").Column
y3 = .Find(" Field Velocity (fps)").Column
End With

'temporaly convert date-time in "site info" into date
'ws1.Range("B23:B200").Copy
'ws1.Range("B23:B200").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks
'     :=False, Transpose:=False
ws1.range("B23:B200").NumberFormat = "mm/dd/yy hh:mm"
'find the length of new data and appending point
Dim append_row_fp As Integer
Dim append_date_fp As Date
Dim first_row_fp As Integer
Dim last_row_fp As Integer
append_row_fp = ws2.Cells(Rows.count, "A").End(xlUp).Row
append_date_fp = ws2.Cells(append_row_fp, 1).Value
'first_row_fp = Application.Match(append_date_fp, ws1.Range("B23:B200")) + 23
first_row_fp = ws1.range("B23:B200").Find(Format(append_date_fp, "mm/dd/yy hh:mm"), LookIn:=xlValues, lookat:=xlWhole).Row
last_row_fp = ws1.Cells(Rows.count, "C").End(xlUp).Row

'copy new date
If first_row_fp <> 0 And last_row_fp <> 0 Then
    For j = first_row_fp To last_row_fp
        ws2.Cells(append_row_fp + j - first_row_fp + 1, 1) = ws1.Cells(j, y0).Value
        ws2.Cells(append_row_fp + j - first_row_fp + 1, 2) = ws1.Cells(j, y1).Value
        ws2.Cells(append_row_fp + j - first_row_fp + 1, 3) = ws1.Cells(j, y2).Value
        ws2.Cells(append_row_fp + j - first_row_fp + 1, 4) = ws1.Cells(j, y3).Value
    DoEvents
    Next j
Else
    MsgBox ("can't find first_row_fp")
    'Exit Sub
End If

'adjust windows for visual check
ws1.Activate
With ActiveWindow
    .ScrollRow = j - 10
End With

ws2.Activate
With ActiveWindow
    .Width = 400
    .Height = 400
    .Top = 300
    .Left = 0
    .ScrollRow = append_row_fp + j - first_row_fp - 10
    .ScrollColumn = 1
End With
ws2.Columns(1).ColumnWidth = 18

'Part III: combined_QAQC (ws1b->ws3)
Workbooks.Open fileName:=target2
Set wb3 = Workbooks(Dir(target2))
Set ws3 = wb3.Sheets(1)

'adjust windows for visual check
ws3.Activate
With ActiveWindow
    .Width = 400
    .Height = 400
    .Top = 700
    .Left = 0
    .ScrollColumn = 1
End With
ws3.Columns(1).ColumnWidth = 18

'find the appending point
Dim append_row_QC As Long
Dim append_date_QC As Date
append_row_QC = ws3.Cells(Rows.count, "A").End(xlUp).Row
append_date_QC = CDate(ws3.Cells(append_row_QC, 1).Value + TimeValue("0:15"))

Do While append_date_QC < #10/1/2014#
    Call QC_pull(QA_sheet_Q3, ws3, append_date_QC, append_row_QC)
    append_row_QC = ws3.Cells(Rows.count, "A").End(xlUp).Row
    append_date_QC = CDate(ws3.Cells(append_row_QC, 1).Value) + TimeValue("0:15")
Loop
    
    Call QC_pull(QA_sheet, ws3, append_date_QC, append_row_QC)

'adjust windows for visual check
ws3.Activate
With ActiveWindow
    .ScrollRow = Cells(Rows.count, 1).End(xlUp).Row - 10
End With

'update input.csv (step 5)
Workbooks.Open fileName:=target3
Cells(3, 2).Value = Dir(target1)
Cells(2, 2).Value = Dir(target2)
Cells(6, 2).Value = ""
Cells(7, 2).Value = ""
'adjust windows for visual check
Workbooks("input.csv").Activate
With ActiveWindow
    .Width = 800
    .Height = 200
    .Top = 1100
    .Left = 0
    .ScrollRow = 1
    .ScrollColumn = 1
End With

endTime = ws3.Cells(Rows.count, 1).End(xlUp).Value
'save updated file
wb3.Close savechanges:=True
wb2.Close savechanges:=True
Workbooks(Dir(QA_sheet)).Close savechanges:=False
Workbooks("input.csv").Close savechanges:=True
Workbooks(Dir(QA_sheet_Q3)).Close savechanges:=False

'add execute time and notes
ws0.Cells(i, 2) = ws0.Cells(i, 2).Value & " " & Format(Date, "dd-mmm")
ws0.Cells(i, 6) = ws0.Cells(i, 6).Value & " " & "BigPicture done up to " _
    & Format(endTime, "yyyy/mm/dd")
wb0.Save
'give some cpu time to the computer so it won't freeze itself
DoEvents

Next i

Application.DisplayAlerts = False
Application.ScreenUpdating = True
 
End Sub
'open Q4 QA sheet for the site
'find the column number of date time, usually column A


Sub CreateTab()

'Hao Zhang @ 2015.1.23
' this is an abandoned part from BigPictures()
' judge if a new tab needs to be created
' there is a bug that do not work well when there are two BigP tab for the same month



'rename the tab with current month-year
ws0name = Format(Now(), "mmmyy") & "BigP"

'add a new tab if current month is not exist, otherwise, ask if overwriting current month tab, if not, create a new tab
'''there is a bug here, no new tab will be added anyhow
Dim ws As Worksheet
For Each ws In wb0.Worksheets
    If ws0name = ws.Name Then
        FileExists = True
        Exit For
    Else
        FileExists = False
    End If
Next ws
    
If FileExists = True Then
    a = MsgBox("There is a tab for the current month already, do you want to overwrite it?", vbYesNoCancel, ws0name & " Exists")
    If a = 6 Then
       Set ws0 = wb0.Sheets(ws0name)
    ElseIf a = 7 Then
    '' copy current sites info from previous month tab
        wb0.Worksheets(ws0name).Copy after:=Sheets(wb0.Worksheets.count)
        Set ws0 = wb0.Sheets(wb0.Worksheets.count)
        ws0.Name = ws0name & "_new"
    ''clear old date and comments
        range("B2", "B" & range("C2").End(xlDown).Row).Clear
        range("f2", "f" & range("C2").End(xlDown).Row).Clear
        range("g2", "g" & range("C2").End(xlDown).Row).Clear
    Else
       Exit Sub
    End If
Else
    '' copy current sites info from previous month tab
    wb0.Worksheets("Dec14BigP").Copy after:=Sheets(wb0.Worksheets.count)
    Set ws0 = wb0.Sheets(wb0.Worksheets.count)
    ws0.Name = ws0name
     ''clear old date and comments
    range("B2", "B" & range("C2").End(xlDown).Row).Clear
    range("f2", "f" & range("C2").End(xlDown).Row).Clear
    range("g2", "g" & range("C2").End(xlDown).Row).Clear
End If

End Sub

Function QC_pull(src, tgt As Worksheet, appDate As Date, appRow As Long)
'src=ws1b (Q3-14, Q4-14) , tgt=target2 (QC file)

Dim first_row_QC As Long
Dim last_row_QC As Long
Dim wb1bb As Workbook
Dim ws1bb As Worksheet

Workbooks.Open fileName:=src
Set wb1bb = Workbooks(Dir(src))
Set ws1bb = wb1bb.Sheets("Flow Data")

'adjust windows for visual check
ws1bb.Activate
With ActiveWindow
    .Width = 400
    .Height = 800
    .Top = 300
    .Left = 400
    .ScrollColumn = 2
End With

With ws1bb.range("A12:AE13")
z0 = .Find("DateTime").Column
z1 = .Find("Level 1").Column
z2 = .Find("Vel 1").Column
z3 = .Find("Flow 1").Column
z4 = .Find("Corrected Flow").Column
z5 = .Find("Corrected Level").Column
End With

'find the row number of first and last entry in QA sheet

'ws1bb.Range("A1:A8846").Copy
'ws1bb.Range("A1:A8846").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
ws1bb.range("A1:A8846").NumberFormat = "mm/dd/yyyy hh:mm:ss"
first_row_QC = ws1bb.range("A:A").Find(Format(appDate, "mm/dd/yyyy hh:mm:ss"), , xlValues, xlWhole).Row
last_row_QC = ws1bb.Cells(ws1bb.Rows.count, "B").End(xlUp).Row
If first_row_QC <> 0 And last_row_QC <> 0 Then
'copy new date
    For k = first_row_QC To last_row_QC Step 1
        tgt.Cells(appRow + k - first_row_QC + 1, 1) = Format(CDbl(ws1bb.Cells(k, z0).Value), "mm/d/yyyy hh:mm")
        tgt.Cells(appRow + k - first_row_QC + 1, 2) = ws1bb.Cells(k, z1).Value
        tgt.Cells(appRow + k - first_row_QC + 1, 3) = ws1bb.Cells(k, z2).Value
        tgt.Cells(appRow + k - first_row_QC + 1, 4) = ws1bb.Cells(k, z3).Value
        tgt.Cells(appRow + k - first_row_QC + 1, 5) = ws1bb.Cells(k, z4).Value
        tgt.Cells(appRow + k - first_row_QC + 1, 6) = ws1bb.Cells(k, z5).Value
    Next k
Else
   MsgBox ("can't find first_row_QC")
   Exit Function
End If
'release the object
Set wb1bb = Nothing
Set ws1bb = Nothing

End Function

Sub BP_file_handling()
'part of Sub BigPicture(), handles folder creation, file move& copy & paste & rename

Dim sites As String
Dim sFound1 As String
Dim sFound2 As String
Dim source1 As String
Dim source2 As String
Dim target1 As String
Dim target2 As String
Static i As Integer
' this is VERY useful to set the default value of a static variable
If i = 0 Then
i = 18
End If

'open the index page if not opened already
'Workbooks.Open FileName:=mypath & Log
Workbooks("QA Logbook.xlsm").Sheets("log").Select
'get site name from QA Logbook

'start the loop that go through each site
sites = Cells(i, 2).Value
'create a sub-folder inside BigPicture with siteID as the folder name
'BigPicture folder must be existed already
'#flow-control needed to avoid errors when folders already existed
MyPath = "C:\Users\hao.zhang\Desktop\BigPicture\" & sites & "\"
If Dir(MyPath, vbDirectory) = Empty Then
MkDir ("C:\Users\hao.zhang\Desktop\BigPicture\" & sites)
End If

'find the first file that partially matches the source file name
'only the filename with extension is assigned to sFound1 and sFound2
' * is used as a wildcard

sFound1 = Dir(Cells(i, 5).Value & "\QAQC\BigPicture\" & sites & "_Combined_field_points*.csv")
sFound2 = Dir(Cells(i, 5).Value & "\QAQC\BigPicture\" & sites & "_CombinedQAQC*.csv")
'set the full path of source and target files
target1 = MyPath & sites & "_Combined_field_points_141121_HZ.csv"
target2 = MyPath & sites & "_CombinedQAQC_141121_HZ.csv"
target3 = MyPath & "input.csv"
target4 = MyPath & "Run_BigPicture.bat"
source1 = Cells(i, 5).Value & "\QAQC\BigPicture\" & sFound1
source2 = Cells(i, 5).Value & "\QAQC\BigPicture\" & sFound2
source3 = Cells(i, 5).Value & "\QAQC\BigPicture\input.csv"
'use the updated bat file from local harddrive instead (changed R version, server name)
source4 = "C:\Users\hao.zhang\Desktop\BigPicture\Run_BigPicture.bat"
bk1 = Cells(i, 5).Value & "\QAQC\BigPicture\bk\" & sFound1
bk2 = Cells(i, 5).Value & "\QAQC\BigPicture\bk\" & sFound2
QA_sheet = Cells(i, 5).Value & "\QAQC\" & Cells(i, 2) & " (Q4-14).xlsm"
'copy Combined_FP from server to local hard drive
If sFound1 <> "" Then
FileCopy source1, target1
'move Combined_FP to bk folder
Name source1 As bk1
End If

'copy Combined_FP from server to local hard drive
If sFound2 <> "" Then
FileCopy source2, target2
'move Combined_FP to bk folder
Name source2 As bk2
End If
'copy input from server to local hard drive
FileCopy source3, target3
'copy Run_bigPicture from server to local hard drive
FileCopy source4, target4


If i = range("B1").End(xlDown).Row Then
rr = MsgBox("that the last one, start from beginning? ", vbYesNo)
    If rr = vbYes Then
        i = 1
    End If
Else
rr = MsgBox(sites & " Done. Move to next site?", vbYesNo)
    If rr = vbYes Then
    i = i + 1
    End If
End If
End Sub

Sub BP_file_open()
'part of Sub BigPicture(),open source and target files for manual copy and paste
Application.ScreenUpdating = False
i = 21  'need to adjust this number every time
Workbooks("QA Logbook.xlsm").Activate
Sheets("log").Select
sites = Cells(i, 2).Value
MyPath = "C:\Users\hao.zhang\Desktop\BigPicture\" & sites & "\"

sFound1 = Dir(Cells(i, 5).Value & "\QAQC\BigPicture\" & sites & "_Combined_field_points*.csv")
sFound2 = Dir(Cells(i, 5).Value & "\QAQC\BigPicture\" & sites & "_CombinedQAQC*.csv")
'set the full path of source and target files
target1 = MyPath & sites & "_Combined_field_points_141121_HZ.csv"
target2 = MyPath & sites & "_CombinedQAQC_141121_HZ.csv"
target3 = MyPath & "input.csv"
target4 = MyPath & "Run_BigPicture.bat"
source1 = Cells(i, 5).Value & "\QAQC\BigPicture\" & sFound1
source2 = Cells(i, 5).Value & "\QAQC\BigPicture\" & sFound2
source3 = Cells(i, 5).Value & "\QAQC\BigPicture\input.csv"
'use the updated bat file from local harddrive instead (changed R version, server name)
source4 = "C:\Users\hao.zhang\Desktop\BigPicture\Run_BigPicture.bat"
bk1 = Cells(i, 5).Value & "\QAQC\BigPicture\bk\" & sFound1
bk2 = Cells(i, 5).Value & "\QAQC\BigPicture\bk\" & sFound2
QA_sheet = Cells(i, 5).Value & "\QAQC\" & Cells(i, 2) & " (Q4-14).xlsm"
'copy Combined_FP from server to local hard drive

Workbooks.Open fileName:=QA_sheet
With ActiveWindow
        .Width = 800
        .Height = 1100
        .Top = 0
        .Left = 0
    End With
    Sheets("Site Info").Select
    range("E23").End(xlDown).Select
    
Workbooks.Open fileName:=target1
 Columns("A:A").ColumnWidth = 18
 With ActiveWindow
        .Width = 350
        .Height = 1100
        .Top = 0
        .Left = 800
    End With
    range("A1").End(xlDown).Select
    
Workbooks.Open fileName:=target2
 Columns("A:A").ColumnWidth = 18
 With ActiveWindow
        .Width = 450
        .Height = 1100
        .Top = 0
        .Left = 1150
    End With
    range("A1").End(xlDown).Select
    
Workbooks.Open fileName:=target3
    With ActiveWindow
        .Left = 0
        .Top = 0
        .Width = 800
        .Height = 300
    End With
    'update input.csv (step 5)
Cells(2, 2).Value = Dir(target2)
Cells(3, 2).Value = Dir(target1)
Application.DisplayAlerts = False
ActiveWorkbook.Save
End Sub

Sub fp_QA_flip()
'2015.1.2 @Hao Zhang
'this is a patch for the error in input.csv that fp and QA files were fliped
Set ws0 = Workbooks("QA Logbook.xlsm").Sheets("Jan15BigP")

For i = 2 To 62
filePath = "C:\Users\hao.zhang\Desktop\BigPicture\" & ws0.Cells(i, 3).Value & "\input.csv"
Workbooks.Open fileName:=filePath
Set ws1 = Workbooks(Dir(filePath))
t = ActiveSheet.Cells(2, 2).Value
ActiveSheet.Cells(2, 2).Value = ActiveSheet.Cells(3, 2).Value
ActiveSheet.Cells(3, 2).Value = t
ws1.Close savechanges:=True
Next i

End Sub

Sub bat_exe()
'2015.1.2 @Hao Zhang
'a trial run for executing Run_BigPicture.bat from VBA
Set ws0 = Workbooks("QA Logbook.xlsm").Sheets("Jan15BigP")
On Error Resume Next
For i = 2 To 62
batpath = "C:\Users\hao.zhang\Desktop\BigPicture\" & Cells(i, 3).Value & "\Run_BigPicture.bat"
Shell "cmd.exe /k" & batpath
Next i

End Sub


Sub file_upload()
'2015.1.5 @Hao Zhang
'1. move existing input files into bk
'2. delete existing big_picture folder
'3. copy new input files into server
'4. copy new big_picture folder into server
'5. loop

Dim ws As Worksheet
Set ws = Workbooks("QA Logbook.xlsm").ActiveSheet
Set objfso = CreateObject("Scripting.Filesystemobject")

'''the loop is temporarly halted'''''
'For i = 2 To ws.Range("F2").End(xlDown).Row
For i = 15 To 53
src = "C:\Users\hao.zhang\Desktop\BigPicture\" & ws.Cells(i, 3)
tgt_1 = ws.Cells(i, 4) & "QAQC\BigPicture"
tgt_2 = ws.Cells(i, 4) & "QAQC\Big Picture"

If objfso.FolderExists(tgt_1) <> 0 Then
tgt = tgt_1
ElseIf objfso.FolderExists(tgt_2) <> 0 Then
tgt = tgt_2
Else
objfso.CreateFolder tgt_1
tgt = tgt_1
End If

tgt_bk = tgt & "\bk\"
QA_old = Dir(tgt & "\*QAQC*.csv")
fp_old = Dir(tgt & "\*field_points*.csv")
QA_new = Dir(src & "\*QAQC*.csv")
fp_new = Dir(src & "\*field_points*.csv")

If QA_old <> "" Then
    If Dir(tgt_bk & QA_old) <> "" Then
        If Dir(tgt_bk & QA_old) <> QA_old Then
        Name tgt & "\" & QA_old As tgt_bk & QA_old
        End If
    ElseIf objfso.FolderExists(tgt_bk) <> 0 Then
    Name tgt & "\" & QA_old As tgt_bk & QA_old
    Else
    objfso.CreateFolder (tgt_bk)
    Name tgt & "\" & QA_old As tgt_bk & QA_old
    End If
End If

If fp_old <> "" Then
    If Dir(tgt_bk & fp_old) <> "" Then
        If Dir(tgt_bk & fp_old) <> fp_old Then
        Name tgt & "\" & fp_old As tgt_bk & fp_old
        End If
    Else
    Name tgt & "\" & fp_old As tgt_bk & fp_old
    End If
End If

objfso.CopyFolder src, tgt, True
Cells(i, 6).Value = Cells(i, 6).Value & ", files uploaded."
Application.Wait (Now() + TimeValue("0:0:2"))
  
DoEvents

Next i

End Sub

Sub xlsChart_ppt()
'Hao Zhang @ 2015.1.21
'Auto-generates the 4 ppt files that contain monthly TS charts from QAQC sheet, and big picture TS plots, referring to the yellow sheet and QA Logbook for good/bad sites
'need to add "Microsoft PowerPoint 14.0 Object library" to the references first
'specify the month and year
Const Month = 12
Const Year = 2014

'declare reference file
Dim yellowWb As Workbook
Dim siteWs As Worksheet
Set yellowWb = Workbooks("Flow Monitoring % Recovery_" & MonthName(Month) & "-HZ.xlsx")
Set siteWs = Workbooks("QA Logbook.xlsm").Worksheets("currentSites")

''create the four ppt files
'Declare the needed variables
Dim newPP As PowerPoint.Application
Dim currentSlide As PowerPoint.Slide
Dim iChart As Excel.Chart
'Check if PowerPoint is active
On Error Resume Next
Set newPP = GetObject(, "PowerPoint.Application")
On Error GoTo 0
'Open PowerPoint if not active
If newPP Is Nothing Then
Set newPP = New PowerPoint.Application
End If
'Create new presentation in PowerPoint
'If newPP.Presentations.Count = 0 Then
newPP.Presentations.Add
'End If
'save the file to the destination folder
newPP.ActivePresentation.SaveAs "C:\Users\hao.zhang\Desktop\" & MonthName(Month) & "_QAQC_Good.pptx"

'Add the cover page slide in PowerPoint for each Excel chart
newPP.ActivePresentation.Slides.Add newPP.ActivePresentation.Slides.count + 1, ppLayoutTitleOnly
newPP.ActiveWindow.View.GotoSlide newPP.ActivePresentation.Slides.count
Set currentSlide = newPP.ActivePresentation.Slides(newPP.ActivePresentation.Slides.count)
'adjust the textbox layout
With currentSlide.Shapes(1)
        .TextFrame.TextRange.Text = MonthName(Month) & " " & Year & " Data" & Chr(13) & "QAQC Good Sites"
        .Width = 620
        .Height = 200
        .Left = 50
        .Top = 100
End With
'Add the RDII title slide
newPP.ActivePresentation.Slides.Add newPP.ActivePresentation.Slides.count + 1, ppLayoutTitleOnly
newPP.ActiveWindow.View.GotoSlide newPP.ActivePresentation.Slides.count
Set currentSlide = newPP.ActivePresentation.Slides(newPP.ActivePresentation.Slides.count)

With currentSlide.Shapes(1)
        .TextFrame.TextRange.Text = "RDII/Separate Sites"
        .Width = 620
        .Height = 200
        .Left = 50
        .Top = 100
        
End With

'Display the PowerPoint presentation
newPP.visible = True


'judge if a site is good or bad
For iRow = 7 To 27
If yellowWb.Sheets(1).Cells(iRow, "K").Font.Color = vbRed Then
    GOB = "Bad"
    tgtQC = MonthName(Month) & "_QAQC_Bad.pptx"
    tgtBP = MonthName(Month) & "_BigPic_Bad.pptx"
Else
    GOB = "Good"
    tgtQC = MonthName(Month) & "_QAQC_Good.pptx"
    tgtBP = MonthName(Month) & "_BigPic_Good.pptx"
End If
Next iRow
'opens the file
Dim ws_src As Worksheet
Set ws_src = Workbooks("QA Logbook.xlsm").Worksheets("currentSites")
Dim C As range
For Each C In ws_src.range("A2:A65")
    If C.Value Like Cells(iRow, 1).Value & "*" Then
        Workbooks.Open fileName:=C.Hyperlinks(1).Address
        Set srcTS = ActiveWorkbook.Charts(Format(MonthName(Month), "mmm") & " TS")
        Set srcTSCorr = ActiveWorkbook.Charts(Format(MonthName(Month), "mmm") & " TS CORR")
        Exit For
    End If


'call the functioln to export chart
'xlsToPPT srcTS, tgtQC
'xlsToPPT srcTSCorr, tgtQC
'pngToPPT srcPic, tgtBigP
'pngToPPT srcPicCorr, tgtBigP

Next C
'31 to 68
'72 to 73

'retrieve the file

End Sub

Function xlsToPPT(srcFile, Mon, tgtFile)

'ref: http://smallbusiness.chron.com/copy-chart-excel-powerpoint-vba-40834.html
'export charts from excel to powerpoint

'Locate Excel charts to paste into the new PowerPoint presentation
'Add a new slide in PowerPoint for each Excel chart
newPP.ActivePresentation.Slides.Add newPP.ActivePresentation.Slides.count + 1, ppLayoutText
newPP.ActiveWindow.View.GotoSlide newPP.ActivePresentation.Slides.count
Set currentSlide = newPP.ActivePresentation.Slides(newPP.ActivePresentation.Slides.count)
'Copy each Excel chart and paste it into PowerPoint as an Metafile image
Xchart.Select
ActiveChart.ChartArea.Copy
currentSlide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
'Copy and paste chart title as the slide title in PowerPoint
currentSlide.Shapes(1).TextFrame.TextRange.Text = Cht.Chart.ChartTitle.Text
'Adjust the slide position for each chart slide in PowerPoint.
'Note that you can adjust the values to position the chart on the slide to your liking
newPP.ActiveWindow.Selection.ShapeRange.Left = 25
newPP.ActiveWindow.Selection.ShapeRange.Top = 150
currentSlide.Shapes(2).Width = 250
currentSlide.Shapes(2).Left = 500


AppActivate ("Microsoft PowerPoint")
Set currentSlide = Nothing
Set newPP = Nothing
End Function
