Attribute VB_Name = "TestRun"
Sub range()
RDII_Row = 6
DCIA_Row = 30
SWM_Row = 72
    For iRow = RDII_Row To SWM_Row + 2
        'Add plots
        If (iRow < 20 And iRow > 10) Or (30 < iRow < 40) Then
            Debug.Print iRow
        End If
    Next


End Sub
Private Sub AllButton_Click()
'select or deselect all items in listbox

Dim ListLength As Integer
Dim counter As Integer

ListLength = ListBox1.ListCount

For counter = 0 To ListLength - 1
ListBox1.Selected(counter) = True 'or False to deselect
Next

End Sub


Sub Macro2()
'
' Macro2 Macro
'change text color of a cell

    range("K34").Select
    With Selection.Font
        .Color = -16776961 ' or vbRed or &hFF
        .TintAndShade = 0
    End With
    
    
    'judge if a site is good or bad
    With Application.FindFormat.Font
        .Subscript = False
        .Color = 255
        .TintAndShade = 0
    End With
    Cells.Find(What:="*", after:=ActiveCell, LookIn:=xlValues, lookat:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=True).Activate
    Cells.FindNext(after:=ActiveCell).Activate

End Sub



Sub GetDataFromClosedWorkbook()
'http://www.exceltip.com/files-workbook-and-worksheets-in-vba/read-information-from-a-closed-workbook-using-vba-in-microsoft-excel.html
Dim wb As Workbook
    Application.ScreenUpdating = False ' turn off the screen updating
    Set wb = Workbooks.Open("C:\Foldername\Filename.xls", True, True)
    ' open the source workbook, read only
    With ThisWorkbook.Worksheets("TargetSheetName")
        ' read data from the source workbook
        .range("A10").Formula = wb.Worksheets("SourceSheetName").range("A10").Formula
        .range("A11").Formula = wb.Worksheets("SourceSheetName").range("A20").Formula
        .range("A12").Formula = wb.Worksheets("SourceSheetName").range("A30").Formula
        .range("A13").Formula = wb.Worksheets("SourceSheetName").range("A40").Formula
    End With
    wb.Close False ' close the source workbook without saving any changes
    Set wb = Nothing ' free memory
    Application.ScreenUpdating = True ' turn on the screen updating
End Sub


Sub ReadDataFromAllWorkbooksInFolder()
'http://www.exceltip.com/files-workbook-and-worksheets-in-vba/read-information-from-a-closed-workbook-using-vba-in-microsoft-excel.html
Dim FolderName As String, wbName As String, r As Long, cValue As Variant
Dim wbList() As String, wbCount As Integer, i As Integer
    FolderName = "C:\Foldername"
    ' create list of workbooks in foldername
    wbCount = 0
    wbName = Dir(FolderName & "\" & "*.xls")
    While wbName <> ""
        wbCount = wbCount + 1
        ReDim Preserve wbList(1 To wbCount)
        wbList(wbCount) = wbName
        wbName = Dir
    Wend
    If wbCount = 0 Then Exit Sub
    ' get values from each workbook
    r = 0
    Workbooks.Add
    For i = 1 To wbCount
        r = r + 1
        cValue = GetInfoFromClosedFile(FolderName, wbList(i), "Sheet1", "A1")
        Cells(r, 1).Formula = wbList(i)
        Cells(r, 2).Formula = cValue
    Next i
End Sub

Private Function GetInfoFromClosedFile(ByVal wbPath As String, _
    wbName As String, wsName As String, cellRef As String) As Variant
Dim arg As String
    GetInfoFromClosedFile = ""
    If Right(wbPath, 1) <> "\" Then wbPath = wbPath & "\"
    If Dir(wbPath & "\" & wbName) = "" Then Exit Function
    arg = "'" & wbPath & "[" & wbName & "]" & _
        wsName & "'!" & range(cellRef).Address(True, True, xlR1C1)
    On Error Resume Next
    GetInfoFromClosedFile = ExecuteExcel4Macro(arg)
End Function

Function IsWorkBookOpen_bk(fileName As String)
'check if a workbook is open (from stackoverflow.com)
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
'This is an example of opening a file directory dialog


Public Sub LogReader()
    Dim Pos As Long
    Dim Dialog As Office.FileDialog
    Set Dialog = Application.FileDialog(msoFileDialogFilePicker)

    With Dialog
        .AllowMultiSelect = True
        .ButtonName = "C&onvert"
        .Filters.Clear
        .Filters.Add "Log Files", "*.log", 1
        .title = "Convert Logs to Excel Files"
        .InitialFileName = "M:\ForHao\"
        .InitialView = msoFileDialogViewList

        If .Show Then
            For Pos = 1 To .SelectedItems.count
                'LogRead .SelectedItems.Item(Pos) ' process each file
            Next
        End If
    End With
End Sub
Sub check_exist()
' this is a segment of code that checks if a file exists
If Len(Dir("c:\Instructions.doc")) = 0 Then
   MsgBox "This file does NOT exist."
Else
   MsgBox "This file does exist."
End If

End Sub

Sub find_file()
' AlphaFrog @ 2014.12.23
' http://www.mrexcel.com/forum/excel-questions/628704-function-folder-subfolder-if-file-exist.html
''      NOT WORKING YET
Dim directory As String
Dim fileName As String
Dim i As Integer
Application.ScreenUpdating = False
'directory = "M:\Modeling\Data\Temporary Monitors\Flow Monitoring\Flow Monitoring by ManholeID\"
directory = "C:\Users\hao.zhang\Desktop\20141223_HZ\"
fileName = "*_HZ.sdb"
Do While fileName <> ""
i = i + 1
ActiveSheet.Cells(i, 1) = FindFile(directory, fileName)

'CTRL+BREAK to stop a loop
Loop
Application.ScreenUpdating = True
End Sub

Function FindFile(ByVal strPath As String, ByVal strFile As String) As String
' AlphaFrog @ 2014.12.23
' http://www.mrexcel.com/forum/excel-questions/628704-function-folder-subfolder-if-file-exist.html
    Dim fsoSubfolder As Object
    
    FindFile = ""  'Default value if file not found
    
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    On Error Resume Next
    If Len(Dir(strPath & strFile)) Then
        ' Search for file in current folder
        FindFile = strPath & Dir(strPath & strFile)    'Path and file name
        'FindFile = strPath                              'Path only
    Else
        'Search sub folders
        For Each fsoSubfolder In CreateObject("Scripting.FileSystemObject").GetFolder(strPath).SubFolders
            FindFile = FindFile(fsoSubfolder.path, strFile)
            If FindFile <> "0" Then Exit For
        Next fsoSubfolder
    End If

End Function

Sub Retrieve_File_listing()
'Ankit Kaul @ 2014.12.23
'http://www.exceltrick.com/formulas_macros/vba-dir-function/Dim strFldrList() As String
'enlists all the files inside a current location and its subfolder
''      NOT WORKING YET
Worksheets("Sheet2").Cells(2, 1).Activate
Call Enlist_Directories("C:\Users\hao.zhang\Desktop\20141223_HZ ", 1)
End Sub


Public Sub Enlist_Directories(strPath As String, lngSheet As Long)
'Ankit Kaul @ 2014.12.23
'http://www.exceltrick.com/formulas_macros/vba-dir-function/Dim strFldrList() As String
'enlists all the files inside a current location and its subfolder

Dim lngArrayMax, x As Long
lngArrayMax = 0
strFn = Dir(strPath & "*.*", 23)
While strFn <> ""
  If strFn <> "." And strFn <> ".." Then
    If (GetAttr(strPath & strFn) And vbDirectory) = vbDirectory Then
      lngArrayMax = lngArrayMax + 1
    ReDim Preserve strFldrList(lngArrayMax)
      strFldrList(lngArrayMax) = strPath & strFn & "\"
    Else
    ActiveCell.Value = strPath & strFn
    Worksheets(lngSheet).Cells(ActiveCell.Row + 1, 1).Activate
    End If
  End If
  strFn = Dir()
Wend
If lngArrayMax <> 0 Then
  For x = 1 To lngArrayMax
    Call Enlist_Directories(strFldrList(x), lngSheet)
  Next
End If
End Sub

Sub Iterate_Files()
'Ankit Kaul @ 2014.12.23
'http://www.exceltrick.com/formulas_macros/vba-dir-function/Dim strFldrList() As String
'enlists all the files inside a current location and its subfolder
Dim Ctr As Integer
Ctr = 1
path = "C:\Windows\ " ' Path should always contain a '\' at end
File = Dir(path) ' Retrieving the first entry.
Do Until File = "" ' Start the loop.
  ActiveSheet.Cells(Ctr, 1).Value = path & File
  Ctr = Ctr + 1
  File = Dir() ' Getting next entry.
Loop
End Sub

Sub Iterate_Folders()
'Ankit Kaul @ 2014.12.23
'http://www.exceltrick.com/formulas_macros/vba-dir-function/Dim strFldrList() As String
'enlists all the files inside a current location and its subfolder
Dim Ctr As Integer
Ctr = 1
path = "C:\Windows\ " ' Path should always contain a '\' at end
FirstDir = Dir(path, vbDirectory) ' Retrieving the first entry.
Do Until FirstDir = "" ' Start the loop.
  If (GetAttr(path & FirstDir) And vbDirectory) = vbDirectory Then
    ActiveSheet.Cells(Ctr, 1).Value = path & FirstDir
    Ctr = Ctr + 1
  End If
  FirstDir = Dir() ' Getting next entry.
Loop
End Sub

Sub find_file2()    'returns the path of each subfolder and number of files inside
'Ankit Kaul @ 2014.12.24
'http://www.exceltrick.com/formulas_macros/vba-dir-function/Dim strFldrList() As String
'enlists all the files inside a current location and its subfolder
Dim Ctr As Integer
Dim fl As File
Dim Number_of_files As Integer
Number_of_files = 0
Ctr = 3
ActiveSheet.Cells(1, 1) = "Folder Path"
ActiveSheet.Cells(1, 2) = "Files"
Dim Fldr As Folder
Set Fldr = fso.GetFolder("C:\Users\hao.zhang\Desktop\")
For Each fl In Fldr.Files
If fl.Attributes <> 34 Then
Number_of_files = Number_of_files + 1
End If
Next fl
ActiveSheet.Cells(2, 1) = Fldr.path
ActiveSheet.Cells(2, 2) = Number_of_files
If Fldr.SubFolders.count > 1 Then
Recursive_Count Fldr, Ctr
End If
End Sub

Function Recursive_Count(SFolder As Folder, Ctr As Integer)
On Error GoTo ErrorHandler
Dim Number_of_files As Integer
Dim Sub_Fldr As Folder
Number_of_files = 0
For Each Sub_Fldr In SFolder.SubFolders
For Each fl In Sub_Fldr.Files
If fl.Attributes <> 34 Then
Number_of_files = Number_of_files + 1
End If
Next fl
ActiveSheet.Cells(Ctr, 1) = Sub_Fldr.path
ActiveSheet.Cells(Ctr, 2) = Number_of_files
Ctr = Ctr + 1
If Sub_Fldr.SubFolders.count > 0 Then
Recursive_Count Sub_Fldr, Ctr
End If
Next Sub_Fldr
Exit Function
ErrorHandler:
MsgBox "Some Error Occurred at " & Sub_Fldr.path & vbNewLine & "Press OK To continue!"
End Function


***********************append data from excel to csv******************
' http://www.mrexcel.com/forum/excel-questions/454693-export-data-adding-existing-csv-file.html
For r = 1 To 10 'loop through row 1 to 10
   
    For C = 1 To 5 '5 columns wide
    delim = ""
    If C < lcol Then delim = ","
    Data = Data & Cells(r, C) & delim
    Next C
Open myfile For Append As #1
Print #1, Data
Close #1
Data = ""
Next r
 **************************alternative (w/o loop)*****************************
 Sub outputtotextfile()
    Open "C:\test.txt" For Append As #1
    Print #1, Join(Application.Transpose(Application.Transpose(range("A" & Selection.Row).Resize(, 10))), ", ")
    Close #1
End Sub

Sub pastepng()
TSpng = "C:\Users\Hao\Dropbox\WHL-0065\WHL-0065\2014_Deployment\QAQC\BigPicture\big_picture\uncorrected_ts_(All).png"
TSCORRpng = "C:\Users\Hao\Dropbox\WHL-0065\WHL-0065\2014_Deployment\QAQC\BigPicture\big_picture\zoomed_plots\corrected_ts_(All).png"
Dim newppt As New PowerPoint.Application
Dim CurSlide As PowerPoint.Slide
Dim goodPPT As Presentation
Set goodPPT = newppt.Presentations.Add(msoCTrue)
pptName = "C:\BigPicture_Good.pptx"

Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutBlank)
CurSlide.Shapes.AddPicture TSpng, msoFalse, msoTrue, -1, -1, -1, -1
Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutBlank)
CurSlide.Shapes.AddPicture TSCORRpng, msoFalse, msoTrue, 0, 0, -1, -1
goodPPT.SaveAs pptName

End Sub
Private Sub SSOAPquery()

target3 = "C:\Users\Hao\Dropbox\SSOAP\Test20150421.sdb"
siteID = "D45-000010"
DrainArea = 136.23453
rg_start_time = "1/1/2013"
rg_end_time = "1/1/2015"
ini = "HZ"

Dim dbs As DAO.Database
Dim tdf As DAO.TableDef

Set dbs = OpenDatabase(target3)
For Each tdf In dbs.TableDefs
    If Not (tdf.Name Like "*Units" Or tdf.Name Like "Holidays" Or tdf.Name Like "Metadata" Or tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
        dbs.Execute "delete * from " & tdf.Name
    End If
Next

'write queries for each table

qRG = "INSERT INTO Raingauges (RaingaugeID,RaingaugeName,RaingaugeLocationX,RaingaugeLocationY,RainUnitID,TimeStep,StartDateTime,EndDateTime) VALUES (1,'" & siteID & "', 0, 0, 1, 15,#" & rg_start_time & "#,#" & rg_end_time & "#);"
qRC = "INSERT INTO RainConverters (RainConverterID,RainConverterName,RainUnitID,Format,LinesToSkip,MonthColumn,MonthWidth,DayColumn,DayWidth,YearColumn,YearWidth,HourColumn,HourWidth,MinuteColumn,MinuteWidth,RainColumn,RainWidth,CodeColumn,CodeWidth,MilitaryTime,AMPMColumn) VALUES (1,'" & siteID & "', 1, 'CSV',1,1,2,2,2,3,4,4,2,5,2,6,8,0,0,True,7);"
qMt = "INSERT INTO Meters (MeterID, MeterName, StartDateTime, EndDateTime,Timestep, FlowUnitID, Area) VALUES (1,'" & siteID & "',#" & rg_start_time & "#,#" & rg_end_time & "#, 15, 1," & DrainArea & ");"
qFC = "INSERT INTO FlowConverters (FlowConverterID,FlowConverterName,FlowUnitID,Format,LinesToSkip,MonthColumn,MonthWidth,DayColumn,DayWidth,YearColumn,YearWidth,HourColumn,HourWidth,MinuteColumn,MinuteWidth,FlowColumn,FlowWidth,CodeColumn,CodeWidth,MilitaryTime) VALUES (1,'" & siteID & "', 1, 'CSV',1,1,2,2,2,3,4,4,2,5,2,6,8,0,0,True);"
qAn = "INSERT INTO Analyses (AnalysisID,AnalysisName,MeterID,RainGaugeID,BaseFlowRate,MaxDepressionStorage,RateOfReduction,InitialValue,R1,R2,R3,t1,T2,T3,K1,K2,K3,RunningAverageDuration,SundayDWFAdj,MondayDWFAdj,TuesdayDWFAdj,WednesdayDWFAdj,ThursdayDWFAdj,FridayDWFAdj,SaturdayDWFAdj,MaxDepressionStorage2,RateOfReduction2,InitialValue2,MaxDepressionStorage3,RateOfReduction3,InitialValue3) VALUES (1,'" & siteID & "_" & Format(Now(), "YYMMDD") & "_" & ini & "',1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0);"

dbs.Execute qRG
dbs.Execute qRC
dbs.Execute qMt
dbs.Execute qFC
dbs.Execute qAn

Set dbs = Nothing

End Sub
