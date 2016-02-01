Attribute VB_Name = "CurSiteTbl_Update"
'restrict variables and procedures available within project
Option Private Module

Dim MyWb As Workbook
'refer to the CurSitesTbl worksheet
Dim MyWs As Worksheet
Dim DrainWs As Worksheet
Dim SiteCount As Integer
Dim listBox_Col As Integer
Dim siteName_Col As Integer
Dim intvl_Col As Integer
Dim siteFldr_Col As Integer
Dim DrainArea_Col As Integer
Dim startQA_Col As Integer
'refer to the column of the last QA sheet in the CurSitesTbl
Dim endQA_Col As Integer
'refer to the column of the specific Quarter and Year in the CurSitesTbl


Sub TblDef()
'public procedure to load table defination
'refer to the current workbook
Set MyWb = Application.ThisWorkbook
Set MyWs = MyWb.Sheets("CurSitesTbl")
Set DrainWs = MyWb.Sheets("TempFlowMon_Sheds")

SiteCount = MyWs.Cells(MyWs.Rows.count, 1).End(xlUp).Row - 1

With MyWs.Rows(1)
    listBox_Col = .Find("ListBox Item").Column
    siteName_Col = .Find("Site Name").Column
    intvl_Col = .Find("Interval (min)").Column
    siteFldr_Col = .Find("Site folder").Column
    DrainArea_Col = .Find("Drainage Area (Acre)").Column
    'get first and last column of QA sheets of CurSitesTbl
    startQA_Col = .Rows(1).Find("Q1-11").Column
    endQA_Col = MyWs.Cells(1, startQA_Col).End(xlToRight).Column
End With

End Sub

Private Sub UpdFileList()
'Hao Zhang @ 2015.07.15
'This part was moved from Userform 'LaunchPad' because it is better handled in the worksheet
'what it does:
' clear the table contents, import site list from "Temporary Flow Monitoring Install Removed Tracking Sheet.xlsx"
'(user might need to spcify the tracking sheet path), then populate the CurSitesTbl

Call TblDef
'set the inital guess of the tracking sheet, which may be moved or renamed over time
';maybe add a textbox for tracking sheet address
trkshtpath = "M:\Data\Temporary Monitors\Flow Monitoring\Current Sites\Temporary Flow Monitoring Install Removed Tracking Sheet.xlsx"

If fso.FileExists(trkshtpath) = True Then
    If IsWorkBookOpen(fso.GetFileName(trkshtpath)) = False Then
        Set trkshtwb = Workbooks.Open(fileName:=trkshtpath)
    Else
        Set trkshtwb = Workbooks(fso.GetFileName(trkshtpath))
    End If
        Set trkshtws = trkshtwb.Sheets(1)
Else
    a = MsgBox("Tracking sheet cannot be found, please specify:", vbOKCancel)
    If a = vbOK Then
        trkshtpath = GetFile(fso.GetParentFolderName(trkshtpath))
    Else
        MsgBox ("The update process has been terminated.")
        Exit Sub
    End If
End If

'create a backup tab before updating file list, if it is there already, then ignore this operation

If TabExists(Format(Now(), "yymmdd") & "_bk", MyWb) = False Then
    MyWs.Copy after:=MyWb.Sheets(MyWb.Worksheets.count)
    MyWb.Sheets(MyWb.Worksheets.count).Name = Format(Now(), "yymmdd") & "_bk"
End If

'Clear all entries
Heading = MyWs.Rows(1)
MyWs.range(MyWs.UsedRange.Address).Clear
MyWs.Rows(1) = Heading
'get the new site count. Note: since sites starts at row 2, the sitecount = lastRow-1
SiteCount = trkshtws.Cells(trkshtws.Rows.count, 1).End(xlUp).Row - 1
'copy new site list to CurSitesTbl
'MyWs.range(Cells(2, siteName_Col).Address, .Cells(SiteCount, siteName_Col).Address).Value = .range("A2:A" & SiteCount).Value

'fill table by row
For iRow = 2 To SiteCount + 1
    'fill [index]
    MyWs.Cells(iRow, 1) = iRow - 1
    'fill col[Site Name]
    MyWs.Cells(iRow, siteName_Col).Value = trkshtws.Cells(iRow, 1).Value
    
    If trkshtws.Cells(iRow, 3).Value = "Present" Then
        'determine if the site is active or removed; temporally fill in col[Row Number]
        MyWs.Cells(iRow, 2).Value = "Active"
        'fill col[ListBox Item]
        MyWs.Cells(iRow, listBox_Col).Value = MyWs.Cells(iRow, siteName_Col).Value
    Else
        MyWs.Cells(iRow, 2).Value = "Removed"
        MyWs.Cells(iRow, listBox_Col).Value = MyWs.Cells(iRow, siteName_Col).Value & "(Removed)"
    End If
    
    'fill col[Drainage Area]
    Set rngFound = DrainTbl.Columns("G").Find(MyWs.Cells(iRow, siteName_Col).Value, lookat:=xlWhole)
    If Not rngFound Is Nothing Then
        MyWs.Cells(iRow, 5).Value = rngFound.Offset(0, -2).Value
    End If
    
    'fill col[Interval]
    ';need to add a control for 15 vs 2 min
    MyWs.Cells(iRow, 6).Value = 15
    
    'fill col[site folders]
    Dim rootFldr As Folder
    Dim Fldr As Folder
    rootStr = rootPath & MyWs.Cells(iRow, siteName_Col).Value
    If fso.FolderExists(rootStr) = True Then
        Set rootFldr = fso.GetFolder(rootStr)
        'find the site folders and write it into cells(iRow, 6)
        For Each Fldr In rootFldr.SubFolders
            'search the folders in the same level, 1 at a time
            If Fldr.Name = "QAQC" Then
                MyWs.Cells(iRow, siteFldr_Col).Value = Fldr.ParentFolder
                MyWs.Hyperlinks.Add anchor:=MyWs.Cells(iRow, siteFldr_Col), Address:=Fldr.ParentFolder
                Exit For
            End If
            'search the sub folders of the current folder, 1 at a time
            'use recursive to get into the ground level, then return to the upper level
            If Fldr.SubFolders.count > 0 Then
                RecurPath Fldr, iRow
            End If
        Next Fldr
    End If
    
Next iRow

End Sub


Private Sub populate_QApath()
'Hao Zhang @2015.7.15
'populate the QApaths based on the site Name

Call TblDef

'fill table by row
For iRow = 2 To SiteCount + 1
 'loop through columns for each quarter from Q1-2011 to Q4-2017
    If MyWs.Cells(iRow, siteFldr_Col).Value <> "" Then
        For jCol = startQA_Col To endQA_Col
            'clear the cell before input
            MyWs.Cells(iRow, jCol).Clear
            GetQApath MyWs.Cells(iRow, siteName_Col).Value, MyWs.Cells(1, jCol).Value
        Next jCol
    End If
Next iRow

'set wrap text and row height, so characters are not overlaping on blank cell
MyWs.range(Cells(1, siteFldr_Col).Address, Cells(1, endQA_Col).Address).EntireColumn.WrapText = True
MyWs.Rows.RowHeight = 15

End Sub
   
Function GetQApath(siteName, QtrYr)
'Hao Zhang @ 2015.2.1
'returns the full path of the specific QA sheet
''revised 2015.2.5
''On Error GoTo ErrorHandler

Dim fl As File
Dim Fldr As Folder
Dim pathStr As String
Dim rPath As String
Dim rFldr As Folder
Dim ws As Worksheet

'get site info
site_Row = MyWs.Columns(3).Find(siteName).Row
QtrYr_Col = MyWs.Rows(1).Find(QtrYr).Column

siteName = MyWs.Cells(site_Row, siteName_Col).Value
intvl = MyWs.Cells(site_Row, intvl_Col).Value
siteFldr = MyWs.Cells(site_Row, siteFldr_Col).Value

'looking in the CurSitesTbl first, if not exist, do the recursion
If MyWs.Cells(site_Row, QtrYr_Col).Value <> "" Then
GetQApath = MyWs.Cells(site_Row, QtrYr_Col).Value
Exit Function
End If

Set Fldr = fso.GetFolder(rootPath & siteName)
QAQCFldr = MyWs.Cells(site_Row, siteFldr_Col).Value & "\QAQC"
'QAQCFldr = RecurFldr(Fldr)


If QAQCFldr <> "" Then
    'loop through columns for each quarter from Q1-2011 to Q4-2017
    For Each fl In fso.GetFolder(QAQCFldr).Files
        'process the 15min files
        If intvl = 15 Then
            'break search criteria into multiple pieces for maximum robustness in file names
            'aka, extra/lack spaces, different connect symbols can still be found
            If InStr(1, fl.Name, siteName, vbTextCompare) <> 0 And InStr(1, fl.Name, QtrYr, vbTextCompare) <> 0 And InStr(1, fl.Name, "min", vbTextCompare) = 0 And InStr(1, fl.Name, ".xls", vbTextCompare) <> 0 Then
             'If fl.Name Like SiteName Then
                GetQApath = fl.path
                MyWs.Cells(site_Row, QtrYr_Col).Value = fl.path
                MyWs.Hyperlinks.Add anchor:=MyWs.Cells(site_Row, QtrYr_Col), Address:=fl.path
                Exit For
            End If
            ''no need to go deep into subfolders since the QAQC sheets are all in the same level
        'process the 5min and 2min files
        Else
            If InStr(1, fl.Name, siteName, vbTextCompare) <> 0 And InStr(1, fl.Name, QtrYr, vbTextCompare) <> 0 And InStr(1, fl.Name, intvl & "min", vbTextCompare) <> 0 And InStr(1, fl.Name, ".xls", vbTextCompare) <> 0 Then
                GetQApath = fl.path
                MyWs.Cells(site_Row, QtrYr_Col).Value = fl.path
                MyWs.Hyperlinks.Add anchor:=MyWs.Cells(site_Row, QtrYr_Col), Address:=fl.path
                Exit For
            End If
            DoEvents
        End If
        DoEvents
    Next
End If


End Function
Function RecurFldr(Fldr As Folder)
'Hao Zhang @ 2015.1.24
'the recursor, part of UpdFileList_Click()
'revised 2015.2.5

Dim Sub_Fldr As Folder

For Each Sub_Fldr In Fldr.SubFolders
        'If InStr(1, Sub_Fldr.Name, "QAQC", vbTextCompare) <> 0 Then
        If Sub_Fldr.Name = "QAQC" Then
            RecurFldr = Sub_Fldr
            Exit For
        End If
    If Sub_Fldr.SubFolders.count > 0 Then
        RecurFldr = RecurFldr(Sub_Fldr)
    End If
    DoEvents
Next Sub_Fldr

End Function


