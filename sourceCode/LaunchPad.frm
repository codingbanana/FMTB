VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LaunchPad 
   Caption         =   "The LaunchPad"
   ClientHeight    =   9660
   ClientLeft      =   48
   ClientTop       =   16380
   ClientWidth     =   13608
   OleObjectBlob   =   "LaunchPad.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "LaunchPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Hao Zhang @ 2015.1.24
'This project is expecting to provide full-stack data ETL service with minimum user input
'This project is developed and debugged under Office 2010 (ver. 14.0) environment

'This project requires the following references:
'   OLE Automation
'   Microsoft Scripting Runtime
'   Microsoft Office 14.0 Objects Library
'   Microsoft PowerPoint 14.0 Objects Library
'   Microsoft Excel 14.0 Objects Library
'   Microsoft Forms 2.0 Objects Library
'   Microsoft ActiveX Data Objects 6.0 Library
'   Microsoft DAO 3.6 Objects Library

''''''''this part is abandoned because it requires every user to have additional package installed, which is impractical''''''''''''
'http://chandoo.org/wp/2013/11/13/pop-up-calendar-excel-vba/
'The calendar control requires mscomct2.ocx, if it's not in C:\Windows\System32\, it can be downloaded from:
'http://activex.microsoft.com/controls/vb6/mscomct2.cab
'after the file is in place, run cmd:
' > regsvr32 c:\windows\system32\mscomct2.ocx
'Then in VBE, check the box in 'Tools'->'Additional Controls'-> 'Microsoft MonthView Control 6.0 (SP6)'
'Finally, a new icon will show up in the toolbox.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim fso As New FileSystemObject

'***************************************************************************************
'refer to the root path of the flow monitoring data sets
Const rootPath As String = "M:\Data\Temporary Monitors\Flow Monitoring\Flow Monitoring by ManholeID\"
'***************************************************************************************
'refer to the current workbook
Dim MyWb As Workbook
'refer to the CurSitesTbl worksheet
Dim MyWs As Worksheet
'refer to the full path of the current workbook
Dim MyPath As String
'refer to the desktop of the local computer
Dim sPath As String
'refer to the value in the 'ListBox Item' column of CurSitsTbl
Dim ListBox As String
'refer to the value in the 'Site Name' column of CurSitsTbl
Dim siteName As String
'refer to the value in the 'Interval (min)' column of CurSitsTbl
Dim intvl As Variant
'refer to the QA sheet for certain site and QtrYr
Dim QAwb As Workbook
'refer to the full path of the QA sheet
Dim QAsheet As String
'refer to the Quarter and Year, in the format of 'Q4-14'
Dim QtrYr As String
'refer to the Row number of the selected site in the CurSitesTbl
Dim siteRow As Integer
'refer to an array that consists listbox and sitename
Dim CurSites() As Variant
'refer to the total number of sites listed in the CurSitesTbl
Dim SiteCount As Integer
'refer to the column of 'ListBox Item' in the CurSitesTbl
Dim listBox_Col As Integer
'refer to the column of 'site folder' in the CurSitesTbl
Dim siteFldr_Col As Integer
'refer to the column of 'interval (min)' in the CurSitesTbl
Dim intvl_Col As Integer
'refer to the column of 'Site Name' in the CurSitesTbl
Dim siteName_Col As Integer
'refer to the column of the first QA sheet in the CurSitesTbl
Dim startQA_Col As Integer
'refer to the column of the last QA sheet in the CurSitesTbl
Dim endQA_Col As Integer
'refer to the column of the specific Quarter and Year in the CurSitesTbl
Dim QtrYr_Col As Integer
'refer to the column of the 'Drainage Area (Acre)' in the CurSitesTbl
Dim DrainArea_Col As Integer
'refer to the full path of the yellow sheet (QAQC tracking sheet) FOLDER
Dim ylwPath As String
'refer to the full path of the yellow sheet (QAQC tracking sheet) FILE
Dim ylwsheet As String
'refer to the initial of the user
Dim ini As String

'Private Sub UserForm_Activate()
'add a minimize button to userform
''  AddToForm MIN_BOX
'The problem is, the userform cannot be found when minimized
'End Sub

'*************************************************************************
'****************************Input****************************************
'*************************************************************************
Private Sub UserForm_Initialize()

Set MyWb = Application.ThisWorkbook
Set MyWs = MyWb.Sheets("CurSitesTbl")

'MyPath = the directory where the macro is saved
MyPath = MyWb.path
'sPath =  the dirctory of desktop
sPath = Environ("USERPROFILE") & "\Desktop"

SiteCount = MyWs.Cells(MyWs.Rows.count, 1).End(xlUp).Row - 1

With MyWs.Rows(1)
    listBox_Col = .Find("ListBox Item").Column
    siteName_Col = .Find("Site Name").Column
    intvl_Col = .Find("Interval (min)").Column
    siteFldr_Col = .Find("Site folder").Column
    DrainArea_Col = .Find("Drainage Area (Acre)").Column
End With

'Hao Zhang revised @ 2015.2.9
'store site info from CurSiteTbl to an array 'CurSites'

CurSites = MyWs.range(Cells(2, listBox_Col - 1).Address, Cells(SiteCount + 1, siteName_Col - 1).Address).Value

'populate Current Sites List to CurSitesLB based on CurSiteTbl
With CurSitesLB
    .ColumnCount = 2
    'set the row number to be hidden
    .ColumnWidths = "0;"
    .List = CurSites
    ' Listbox.value will be based on the specified BoundColumn
    ' Listbox.text will be based on the specified TextColumn
    .TextColumn = 2
    'when .BoundColumn = 0, listIndex will be returned as the value, starting from 0
    .BoundColumn = 1
End With

With chosenSitesLB
    .ColumnCount = 2
    .ColumnWidths = "0;"
    .TextColumn = 2
    .BoundColumn = 1
    .MultiSelect = fmMultiSelectExtended
End With

With DoneSitesLB
    .ColumnCount = 2
    .ColumnWidths = "0;"
    .TextColumn = 2
    .BoundColumn = 1
    .MultiSelect = fmMultiSelectSingle
End With

'Hao Zhang @ 2015.3.17
'combine QAtempYrCB And QAtempQtrCB
With QAtempQtrYrCB
    'remove the space before each item
    .SelectionMargin = False
    .ColumnCount = 4
    .ColumnWidths = ";0;0;0"
    For yr = 2011 To 2017
        For qtr = 1 To 4
        'col1:Q1-15(quarter-year); col2: 2015(year)
        'col3:1(quarter number); col4:1(starting month)
            .AddItem "Q" & qtr & "-" & Right(yr, 2)
            .List(yr - 2011, 1) = yr
            .List(yr - 2011, 2) = qtr
            .List(yr - 2011, 3) = qtr * 3 - 2
        Next
    Next
    .Text = .List(16)
End With

'combine QAyrCB and QAQtrCB
With QAqtrYrCB
    'remove the space before each item
    .SelectionMargin = False
    .ColumnCount = 4
    .ColumnWidths = ";0;0;0"
    .BoundColumn = 0
    For yr = 2011 To 2017
        For qtr = 1 To 4
            'col1:Q1-15(quarter-year); col2: 2015(year)
            'col3:1(quarter number); col4:1(starting month)
            .AddItem "Q" & qtr & "-" & Right(yr, 2)
            .List(.ListCount - 1, 1) = yr
            .List(.ListCount - 1, 2) = qtr
            .List(.ListCount - 1, 3) = qtr * 3 - 2
        Next
    Next
    .Text = .List(16)
End With

With RGCB
    .SelectionMargin = False
    For N = 1 To 35
        .AddItem N
    Next
End With

With BigPpptYrCB
    .SelectionMargin = False
    For N = 2011 To 2017
        .AddItem N
    Next
End With

With BigPpptMonCB
    .SelectionMargin = False
    For N = 1 To 12
        .AddItem N
    Next
End With

'load intial values from cacheTbl
With MyWb.Sheets("cacheTbl")
    'Hao Zhang @ 2015.3.18
    'QAQC:
    QAtempTB.Text = .Cells(2, 2).Value
    QAtempQtrYrCB.Text = .Cells(3, 2).Value
    QAtempIntvlTB.Text = .Cells(4, 2).Value
    QAqtrYrCB.Text = .Cells(5, 2).Value
    RawDateTB.Text = .Cells(6, 2).Value
    QAiniTB.Text = .Cells(7, 2).Value
    StartTimeTB.Text = .Cells(8, 2).Value
    EndTimeTB.Text = .Cells(9, 2).Value
    'don't load RGCB because it can be determined automatically
    'BigPicture:
    BigPpathTB.Text = .Cells(12, 2).Value
    BigPfullOB.Value = .Cells(13, 2).Value
    BigPappOB.Value = Not .Cells(13, 2).Value
    BigPiniTB.Text = .Cells(14, 2).Value
    BigPpptPathTB.Value = .Cells(15, 2).Value
    BigPpptYrCB.Text = .Cells(16, 2).Value
    BigPpptMonCB.Text = .Cells(17, 2).Value
    'SSOAP:
    SSOAPpathTB.Value = .Cells(19, 2).Value
    SSOAPfullOB.Value = .Cells(20, 2).Value
    SSOAPappOB.Value = Not .Cells(20, 2).Value
    SSOAPinTB = .Cells(21, 2).Value
    sdbPathTB.Text = .Cells(22, 2).Value
    If .Cells(23, 2).Value <> "" Then
        sdbStartDP = .Cells(23, 2).Value
    End If
    If .Cells(24, 2).Value <> "" Then
        sdbEndDP = .Cells(24, 2).Value
    End If
    SiteNameTB.Text = .Cells(25, 2).Value
End With

'get first and last column of QA sheets of CurSitesTbl
startQA_Col = MyWs.Rows(1).Find("Q1-11").Column
endQA_Col = MyWs.Cells(1, startQA_Col).End(xlToRight).Column

End Sub

Private Sub SiteSortBtn_Click()
'Hao Zhang @ 2015.2.10
'Sorts CurSitesLB alphabetically when click the SiteSortBtn
Dim i As Long
Dim j As Long
Dim Temp As Variant

With CurSitesLB
    For j = 0 To .ListCount - 2
        For i = 0 To .ListCount - 2
            If .List(i, 1) > .List(i + 1, 1) Then
                temp0 = .List(i, 0)
                temp1 = .List(i, 1)
                .List(i, 0) = .List(i + 1, 0)
                .List(i, 1) = .List(i + 1, 1)
                .List(i + 1, 0) = temp0
                .List(i + 1, 1) = temp1
            End If
        Next i
    Next j
End With

End Sub

Private Sub ResetBtn_Click()
'Hao Zhang @ 2015.1.29
're-initialize sites list

CurSitesLB.List = CurSites
chosenSitesLB.Clear
DoneSitesLB.Clear

End Sub

Private Sub QAqtrYrCB_afterupdate()
'Hao Zhang @ 2015.3.17
'change the startTimeTB and EndTimeTB for rainfall import after changes in QAqtrYrCB
With QAqtrYrCB
    StartTimeTB.Value = DateSerial(.Column(1), .Column(3), 1)
    EndTimeTB.Value = DateAdd("m", 3, DateSerial(.Column(1), .Column(3), 1))
End With

End Sub

Private Sub QAiniTB_change()
'syncronize initials

BigPiniTB.Text = QAiniTB.Text
SSOAPiniTB.Text = QAiniTB.Text

End Sub

Private Sub AddAllBtn_Click()
'Hao Zhang @ 2015.1.24
'move all sites from current sites to chosen sites

Call ResetBtn_Click

For N = 0 To CurSitesLB.ListCount - 1
'this is the third way to populate a listbox
    chosenSitesLB.AddItem (CurSitesLB.List(N))
    chosenSitesLB.List(chosenSitesLB.ListCount - 1, 1) = CurSitesLB.List(N, 1)
Next N

'pre-select the first item in chosen listbox
If chosenSitesLB.ListIndex <> -1 Then
    chosenSitesLB.Selected(0) = True
End If

'There is a very interesting trick: you need to subtract selected item from the end so the other index won't be changed
For N = CurSitesLB.ListCount - 1 To 0 Step -1
    CurSitesLB.RemoveItem (N)
Next N

End Sub

Private Sub AddBtn_Click()
'Hao Zhang @ 2015.1.24
'move select sites from current sites to chosen sites

For N = 0 To CurSitesLB.ListCount - 1
    If CurSitesLB.Selected(N) = True Then
        chosenSitesLB.AddItem (CurSitesLB.List(N))
        chosenSitesLB.List(chosenSitesLB.ListCount - 1, 1) = CurSitesLB.List(N, 1)
    End If
Next N

'pre-select the first item in chosen listbox
If chosenSitesLB.ListIndex <> -1 Then
    chosenSitesLB.Selected(0) = True
End If

'There is a very interesting trick: you need to subtract selected item from the end so the other index won't be changed
For N = CurSitesLB.ListCount - 1 To 0 Step -1
    If CurSitesLB.Selected(N) = True Then
        CurSitesLB.RemoveItem (N)
    End If
Next N

End Sub

Private Sub SubAllBtn_Click()
'Hao Zhang @ 2015.1.24
'move all sites from chosen sites to current sites

Call ResetBtn_Click

For N = 0 To chosenSitesLB.ListCount - 1
    CurSitesLB.AddItem (chosenSitesLB.List(N))
    CurSitesLB.List(CurSitesLB.ListCount - 1, 1) = chosenSitesLB.List(N, 1)
Next N

'There is a very interesting trick: you need to subtract selected item from the end so the other index won't be changed
For N = chosenSitesLB.ListCount - 1 To 0 Step -1
chosenSitesLB.RemoveItem (N)
Next N


End Sub

Private Sub SubBtn_Click()
'Hao Zhang @ 2015.1.24
'move selected sites from chosen sites to current sites

For N = 0 To chosenSitesLB.ListCount - 1
    If chosenSitesLB.Selected(N) = True Then
        CurSitesLB.AddItem (chosenSitesLB.List(N))
        CurSitesLB.List(CurSitesLB.ListCount - 1, 1) = chosenSitesLB.List(N, 1)
    End If
Next N

For N = chosenSitesLB.ListCount - 1 To 0 Step -1
    If chosenSitesLB.Selected(N) = True Then
        chosenSitesLB.RemoveItem (N)
    End If
Next N
'pre-select the first item in chosen listbox
If chosenSitesLB.ListIndex <> -1 Then
    chosenSitesLB.Selected(0) = True
End If

End Sub

Private Sub CurSitesLB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Hao Zhang @ 2015.2.10
'double click on site to open the site folder
For isite = 0 To CurSitesLB.ListCount - 1
    If CurSitesLB.Selected(isite) = True Then
        siteName = CurSitesLB.List(isite, 1)
        sdbEndDP.Text = CurSitesLB.List(isite, 0)
        Exit For
    End If
Next

siteFldr = MyWs.Cells(siteRow, siteFldr_Col).Value

Shell "C:\WINDOWS\explorer.exe """ & siteFldr & "", vbNormalFocus

End Sub

Private Sub chosenSitesLB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Hao Zhang @ 2015.2.9
'opens the QA sheet when double clicked

For isite = 0 To chosenSitesLB.ListCount - 1
    If chosenSitesLB.Selected(isite) = True Then
        siteName = chosenSitesLB.List(isite, 1)
        siteRow = chosenSitesLB.List(isite, 0)
        Exit For
    End If
Next

QtrYr = QAqtrYrCB.Column(0)
QtrYr_Col = MyWs.Rows(1).Find(QtrYr).Column
QAsheet = GetQApath(siteName, QtrYr)

If QAsheet = "" Then
    MsgBox "The QA sheet has not been created yet."
    Exit Sub
Else
    If IsWorkBookOpen(QAsheet) = True Then
        Workbooks(fso.GetFileName(QAsheet)).Activate
    Else
        Workbooks.Open fileName:=QAsheet
    End If
End If
With Workbooks(fso.GetFileName(QAsheet)).Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth
    .Height = Application.UsableHeight * 0.8
    .Top = 0
    .Left = 0
End With

End Sub
Private Sub MoveDoneSiteBtn_Click()

'Private Sub MoveDoneSite()
'Hao Zhang @ 2015.2.9
'move one selected site from chosen sites LB to done sites LB

For N = chosenSitesLB.ListCount - 1 To 0 Step -1
    If chosenSitesLB.Selected(N) = True Then
        DoneSitesLB.AddItem (chosenSitesLB.List(N))
        DoneSitesLB.List(DoneSitesLB.ListCount - 1, 1) = chosenSitesLB.List(N, 1)
        'only affect the first selected site
        chosenSitesLB.RemoveItem (N)
        Exit For
    End If
Next N

''pre-select the first item in chosen listbox
'If chosenSitesLB.ListIndex <> -1 Then
'    chosenSitesLB.Selected(0) = True
'End If

End Sub

Private Sub DoneSitesLB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'add a feature that puts the site back to chosenSitesLB
'Hao Zhang @ 2015.2.9
'move double-clicked sites from done sites back to the top of chosen sites
'(in case a site needs redo)

For N = 0 To DoneSitesLB.ListCount - 1
    If DoneSitesLB.Selected(N) = True Then
        chosenSitesLB.AddItem DoneSitesLB.List(N), 0
        chosenSitesLB.List(0, 1) = DoneSitesLB.List(N, 1)
        'only affect the first selected site
        DoneSitesLB.RemoveItem (N)
        Exit For
    End If
Next N

'pre-select the first item in chosen listbox
If chosenSitesLB.ListIndex <> -1 Then
    chosenSitesLB.Selected(0) = True
End If
End Sub
Private Sub UpdFileListBtn2_Click()
'Hao Zhang @ 2015.1.24
'fill CurSitesTbl in the active workbook with current sites paths


Dim fl As File
Dim Fldr As Folder
Dim rootStr As String
Dim rootFldr As Folder
Dim iRow As Integer
Dim jCol As Integer

'create a backup tab before updating file list, if it is there already, then ignore this operation

If TabExists(Format(Now(), "yymmdd") & "_bk", MyWb) = False Then
    MyWs.Copy after:=Sheets(MyWb.Worksheets.count)
    Sheets(MyWb.Worksheets.count).Name = Format(Now(), "yymmdd") & "_bk"
End If
'show the CurSitesTbl tab
MyWs.Activate

MyWs.range(Cells(2, siteFldr_Col).Address, "AH" & SiteCount + 1).ClearContents

'get site folders for each site
For iRow = 2 To SiteCount + 1
    'Set Fldr = fso.GetFolder(rootPath & "\" & siteName)
    'MyWs.Cells(iRow, siteFldr_Col) = RecurFldr(Fldr)
    '''there is a [bug] in site USE-0400, where the old development are found instead
    'bypassing the bug
    'If MyWs.Cells(iRow, siteName_Col) <> "USE-0400" Then
        rootStr = rootPath & MyWs.Cells(iRow, siteName_Col).Value
        Set rootFldr = fso.GetFolder(rootStr)
        'find the site folders and write it into cells(iRow, 6)
        For Each Fldr In rootFldr.SubFolders
            'search the folders in the same level, 1 at a time
            'If InStr(1, Fldr.Name, "QAQC", vbTextCompare) <> 0 Then
            If Fldr.Name = "QAQC" Then
                MyWs.Cells(iRow, siteFldr_Col) = Fldr.ParentFolder
                MyWs.Hyperlinks.Add anchor:=MyWs.Cells(iRow, siteFldr_Col), Address:=Fldr.ParentFolder
                Exit For
            End If
            'search the sub folders of the current folder, 1 at a time
            'use recursive to get into the ground level, then return to the upper level
            If Fldr.SubFolders.count > 0 Then
                RecurPath Fldr, iRow
            End If
        Next Fldr
    'End If
Next iRow

For iRow = 2 To SiteCount + 1
    'loop through columns for each quarter from Q1-2011 to Q4-2017
    For jCol = 7 To 34
        GetQApath MyWs.Cells(iRow, siteName_Col).Value, MyWs.Cells(1, jCol).Value
    Next jCol
Next iRow
'        'set sites folders path
'        pathStr = MyWs.Cells(iRow, 4) & "\QAQC"
'        'return the folders within pathStr
'        Set Fldr = fso.GetFolder(pathStr)
'        'clear each target cell first
'        MyWs.Cells(iRow, jCol).ClearContents
'        'process the 15min files
'        If Cells(iRow, 3) = 15 Then
'            For Each fl In Fldr.Files
'                'break search criteria into multiple pieces for maximum robustness in file names
'                'aka, extra/lack spaces, different connect symbols can still be found
'                If InStr(1, fl.Name, MyWs.Cells(iRow, 2), vbTextCompare) <> 0 And InStr(1, fl.Name, MyWs.Cells(1, jCol), vbTextCompare) <> 0 And InStr(1, fl.Name, "min", vbTextCompare) = 0 And InStr(1, fl.Name, ".xls", vbTextCompare) <> 0 Then
'                 'If fl.Name Like SiteName Then
'                    MyWs.Cells(iRow, jCol) = fl.path
'                    Exit For
'                End If
'                ''no need to go deep into subfolders since the QAQC sheets are all in the same level
'                DoEvents
'            Next fl
'        'process the 5min and 2min files
'        Else
'            For Each fl In Fldr.Files
'                If InStr(1, fl.Name, MyWs.Cells(iRow, 2), vbTextCompare) <> 0 And InStr(1, fl.Name, MyWs.Cells(1, jCol), vbTextCompare) <> 0 And InStr(1, fl.Name, Cells(iRow, 3) & "min", vbTextCompare) <> 0 And InStr(1, fl.Name, ".xls", vbTextCompare) <> 0 Then
'                    MyWs.Cells(iRow, jCol) = fl.path
'                    Exit For
'                End If
'                DoEvents
'            Next fl
'        End If
'        DoEvents
'    Next jCol
'    DoEvents
'Next iRow
              
MsgBox "The CurSitesTbl has been updated."

MyWb.Save

'Exit Sub
'this part has not been implemented yet
'ErrorHandler:
'Dim errWs As Worksheet
'Set errWs = MyWb.Sheets("ErrorLog")
'newRow = errWs.Cells(errWs.Rows.Count, 1).End(xlUp).Row + 1
'errWs.Cells(newRow, 1).Value = Now()
'errWs.Cells(newRow, 2).Value = "error on" & siteName
'Resume Next

End Sub
'Function GetSitePath(siteName, keyword)
''Hao Zhang @ 2015.2.13
'
'Dim Fldr As Folder
'
''site_Row = MyWs.Columns(listBox_Col).Find(siteName).Row
''siteName = MyWs.Cells(site_Row, siteName_Col).Value
'Set Fldr = fso.GetFolder(rootPath & siteName)
'GetSitePath = RecFldr(Fldr, keyword)
'''goes into every subfolder and search for folders with name of keyword
''For Each Sub_Fldr In Fldr.SubFolders
''        'If InStr(1, Sub_Fldr.Name, "QAQC", vbTextCompare) <> 0 Then
''        If Sub_Fldr.Name = keyword Then
''            'RecurFldr = Sub_Fldr
''            'load all found QAQC folder path and modified date into a temp array
''            tempArr(i, 0) = Sub_Fldr.path
''            tempArr(i, 1) = Sub_Fldr.DateLastModified
''            i = i + 1
''            'Exit For
''        End If
''    If Sub_Fldr.SubFolders.Count > 0 Then
''        GetSitePath = curfldr(Sub_Fldr, keyword)
''    End If
''    DoEvents
''Next Sub_Fldr
''
'''find the latest QAQC folder (= max in last modified date)
''For i = LBound(tempArr, 1) To UBound(tempArr, 1)
''    If tempArr(i, 1) > tMax Then
''      tMax = Arr(i, 1)
''      Index = i
''    End If
''Next i
''
''GetSitePath = tempArr(Index, 0)
'
'End Function
'
'Function RecFldr(Fldr As Folder, keyword)
''Hao Zhang @ 2015.1.24
''the recursor, part of UpdFileList_Click()
''revised 2015.2.5
'
'Dim Sub_Fldr As Folder
'Dim tempArr(5, 1) As Variant
'Dim i As Integer
'Dim index As Integer
'Dim tMax As Date
'
'For Each Sub_Fldr In Fldr.SubFolders
'    'If InStr(1, Sub_Fldr.Name, "QAQC", vbTextCompare) <> 0 Then
'    If Sub_Fldr.Name = keyword Then
'        'RecurFldr = Sub_Fldr
'        'load all found QAQC folder path and modified date into a temp array
'        tempArr(counter, 0) = Sub_Fldr.path
'        tempArr(counter, 1) = Sub_Fldr.DateLastModified
'        counter = counter + 1
'        'Exit For
'    End If
'    If Sub_Fldr.SubFolders.Count > 0 Then
'    RecFldr = RecFldr(Sub_Fldr, keyword)
'    End If
'    DoEvents
'Next Sub_Fldr
'
''find the latest QAQC folder (= max in last modified date)
'For i = LBound(tempArr, 1) To UBound(tempArr, 1)
'    If tempArr(i, 1) > tMax Then
'      tMax = tempArr(i, 1)
'      index = i
'    End If
'Next i
'
'RecFldr = tempArr(index, 0)
'counter = 0
'End Function

Function RecurPath(Fldr As Folder, ByVal iRow)
'Hao Zhang @ 2015.1.24
'the recursor, part of UpdFileList_Click()

Dim Sub_Fldr As Folder

For Each Sub_Fldr In Fldr.SubFolders
        If InStr(1, Sub_Fldr.Name, "QAQC", vbTextCompare) <> 0 Then
            MyWs.Cells(iRow, siteFldr_Col) = Sub_Fldr.ParentFolder
            MyWs.Hyperlinks.Add anchor:=MyWs.Cells(iRow, siteFldr_Col), Address:=Sub_Fldr.ParentFolder
            Exit For
        End If

    If Sub_Fldr.SubFolders.count > 0 Then
        RecurPath Sub_Fldr, iRow
    End If
    DoEvents
Next Sub_Fldr

End Function
Private Sub UpdFileListBtn_Click()
'Hao Zhang @ 2015.07.14
' clear the table contents, import site list from "Temporary Flow Monitoring Install Removed Tracking Sheet.xlsx"
'(user might need to spcify the tracking sheet path), then populate the CurSitesTbl

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
    
    'loop through columns for each quarter from Q1-2011 to Q4-2017
    If MyWs.Cells(iRow, siteFldr_Col).Value <> "" Then
        For jCol = startQA_Col To endQA_Col
            GetQApath MyWs.Cells(iRow, siteName_Col).Value, MyWs.Cells(1, jCol).Value
        Next jCol
    End If
Next iRow

'set wrap text and row height, so characters are not overlaping on blank cell
MyWs.range(Cells(1, siteFldr_Col).Address, Cells(1, endQA_Col).Address).EntireColumn.WrapText = True
MyWs.Rows.RowHeight = 15


End Sub

Private Sub QAbrowseBtn_Click()
'Hao Zhang @ 2015.1.29
'allow users to type or browse to the source template file
QAtempTB.Text = GetFile(QAtempTB.Text)
End Sub


Private Sub DoneSitesLB_afterupdate()
'Hao Zhang @ 2015.2.11
'save the DoneSitesLB in a hidden table

With MyWb.Worksheets("DoneSitesTbl")
    .UsedRange.Clear
    For isite = 0 To DoneSitesLB.ListCount - 1
        .Cells(isite + 1, 1).Value = DoneSitesLB.List(isite, 0)
        .Cells(isite + 1, 2).Value = DoneSitesLB.List(isite, 1)
    Next
End With
End Sub


'*******************************************************************
'**********************Site Info************************************
'*******************************************************************

'public sub = sub, accessible to all modules in the project
'private sub, accesible only to the current module

Public Sub loadSite()
'Hao Zhang @ 2015.2.10
'load site info: siteName, Row, Col, QAsheet, Raw Data(?), big picture, ssoap...
'get the selected item's name and row number

For isite = 0 To chosenSitesLB.ListCount - 1
    If chosenSitesLB.Selected(isite) = True Then
        siteName = chosenSitesLB.List(isite, 1)
        siteRow = chosenSitesLB.List(isite, 0)
        'only the first selected item is returned
        Exit For
    End If
Next

'if no item in chosenSitesLB is select, then look for opened workbooks that looks like a QA sheet
iCount = 0 'for selected listbox
iFlg = 0 'for opened workbook

'count selected items
For isite = 0 To chosenSitesLB.ListCount - 1
    If chosenSitesLB.Selected(isite) = True Then
        iCount = iCount + 1
    End If
Next
'get site info based on active workbook
If iCount = 0 Then
    For Each wb In Workbooks
        If wb.Name Like "*-*(Q?-??)*.xls*" Then
            Set QAwb = wb
            QAsheet = wb.FullName
            siteName = Left(fso.GetBaseName(QAsheet), InStr(1, fso.GetBaseName(QAsheet), "(", vbTextCompare) - 2)
            QtrYr = Mid(fso.GetBaseName(QAsheet), InStrRev(fso.GetBaseName(QAsheet), "(", , vbTextCompare) + 1, 5)
            If fso.GetBaseName(QAsheet) Like "*min*" Then
                intvl = Mid(fso.GetBaseName(QAsheet), InStr(1, fso.GetBaseName(QAsheet), "min", vbTextCompare) - 1, 1)
                siteName = siteName & "_" & intvl & "min"
            Else
                intvl = 15
            End If
            LB_col = MyWs.Rows(1).Find("ListBox Item").Column
            siteRow = MyWs.Columns(LB_col).Find(siteName).Row
            QtrYr_Col = MyWs.Rows(1).Find(QtrYr).Column
            iFlg = iFlg + 1
        End If
    Next
        If iFlg = 0 Then
            MsgBox "Can't find any selected or opened QA sheet, please select or open some sites and try again"
            End
        ElseIf iFlg > 1 Then
            MsgBox "there are more than one QA sheets, please leave only one sheet open"
            End
        End If
'get site info based on selected site in listbox
Else
    QtrYr = QAqtrYrCB.Column(0)
    QtrYr_Col = MyWs.Rows(1).Find(QtrYr).Column
    QAsheet = GetQApath(siteName, QtrYr)
    'if QA sheet doesn't exist, then create a new QA sheet from a template, and write the path in CurSitesTbl
    If QAsheet = "" Then
        a = MsgBox("The QA sheet is not exist, would you like to create one?" & Chr(13) & "You will be asked to choose a template.", vbYesNo)
        If a = vbYes Then
            'create a new QA sheet from an existing template
            siteFldr_Col = MyWs.Rows(1).Find("Site folder").Column
            intvl_Col = MyWs.Rows(1).Find("Interval (min)").Column
            intvl = MyWs.Cells(siteRow, intvl_Col).Value
            If intvl = 15 Then
                QAsheet = MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\" & siteName & " (" & QtrYr & ").xlsm"
            Else
                QAsheet = MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\" & siteName & " (" & QtrYr & ")_" & intvl & "min.xlsm"
            End If
            tempQAsheet = GetFile(QAtempTB.Value)
            FileCopy tempQAsheet, QAsheet
            'write the path into CurSitesTbl
            MyWs.Cells(siteRow, QtrYr_Col).Value = QAsheet
            MyWs.Hyperlinks.Add anchor:=MyWs.Cells(siteRow, QtrYr_Col), Address:=QAsheet
            Set QAwb = Workbooks.Open(fileName:=QAsheet)
        Else
            MsgBox "Please create the QA sheet manually and try again."
            End
        End If
    'if QA sheet exists, check if the workbook is already opened
    Else
        If IsWorkBookOpen(QAsheet) = True Then
            Set QAwb = Workbooks(fso.GetFileName(QAsheet))
        Else
            Set QAwb = Workbooks.Open(fileName:=QAsheet)
        End If
        intvl_Col = MyWs.Rows(1).Find("Interval (min)").Column
        intvl = MyWs.Cells(siteRow, intvl_Col).Value
    End If
    
End If

'adjust windows for visual check
With QAwb.Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth
    .Height = Application.UsableHeight * 0.8
    .Top = 0
    .Left = 0
End With


End Sub


'**************************************************************************
'****************************QA Template***********************************
'**************************************************************************

Private Sub QAtempGenBtn_Click()
'Hao Zhang @ 2015.1.29
'update QA template to any quarter, code adopted from [template]changeScale()
'can only handle new version of template
Dim wb As Workbook

srcPath = QAtempTB.Text

'enter the time interval in minutes
intvl = QAtempIntvlTB.Text

'use the fso method to get the dirctory
flpath = fso.GetParentFolderName(srcPath) & "\"
srcFile = fso.GetBaseName(srcPath)
srcExt = "." & fso.GetExtensionName(srcPath)
tgtName = "QAQC template (" & QAtempQtrYrCB.Column(0) & ")_" & intvl & "min"
'get an unique file name using function GetUniqueName()
tgtPath = GetUniqueName(flpath & tgtName & srcExt)

'duplicate source file, rename the generated file.
If fso.FileExists(srcPath) = True Then
    'if sourcefile and target file are same, then not duplicate is needed
    If srcPath <> tgtPath Then
    FileCopy srcPath, tgtPath
    End If
Else
    MsgBox "the source template doesn't exist."
    Exit Sub
End If

'Hao Zhang @ 2015.1.30
'opens the new template file
Application.Workbooks.Open fileName:=tgtPath
Set wb = Workbooks(fso.GetFileName(tgtPath))
'stop screen flashing
Application.ScreenUpdating = False

'Part I: worksheets
'set up the StartDate and endDate, endDate will be updated for each month in functions
month1 = DateSerial(mYear, startMonth, 1)
Month2 = DateAdd("m", 1, month1)
Month3 = DateAdd("m", 2, month1)
Month4 = DateAdd("m", 3, month1)
'identify the row number of the begining of each month (Row1,2,3) and the end of the last month(Row4)
Dim Row1 As Long
Dim Row2 As Long
Dim Row3 As Long
Dim Row4 As Long

Row1 = 14
Row2 = Row1 + ((Day(Month2 - 1)) * 60 / intvl * 24)
Row3 = Row2 + ((Day(Month3 - 1)) * 60 / intvl * 24)
Row4 = Row3 + ((Day(Month4 - 1)) * 60 / intvl * 24)

'update the month-year in sheet Site Info
With wb.Sheets("Site Info")
'    Date_Col = .Range("10:25").Find("Date").Column
'    FLvl_Col = .Range("10:25").Find(" Field Level (inches)").Column
'    rng = .Range("10:25").Find("Area (ft2)")
'    Area1_Col = rng.Column
'    Area2_Col = .Range("10:25").FindNext(rng).Column
'    FFlw_Col = .Range("10:25").Find("Field Flow (mgd)").Column
'    MFlw_Col = .Range("10:25").Find("Meter Flow (mgd)").Column
    .range("B2:C13").ClearContents
    .range("F2:J3").ClearContents
    .range("B16:V194").ClearContents
    .range("B16:B194").Formula = "=C16+D16"
    .range("I16:I194").Formula = "=IF(OR(R16=" & Chr(34) & Chr(34) & ",H16=" & Chr(34) & Chr(34) & ")," & Chr(34) & Chr(34) & ",(VLOOKUP(H16,'Area vs. depth table'!A:C,3,TRUE)-VLOOKUP(R16,'Area vs. depth table'!A:C,3,TRUE)))"
    .range("J16:J194").Formula = "=IF(I16=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",I16*K16*0.6463)"
    .range("O16:O194").Formula = "=IF(OR(R16=" & Chr(34) & Chr(34) & ")," & Chr(34) & Chr(34) & ",(VLOOKUP(M16,'Area vs. depth table'!A:C,3,TRUE)-VLOOKUP(M16,'Area vs. depth table'!A:C,3,TRUE)))"
    .range("P16:P194").Formula = "=IF(O16=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",O16*Q16*0.6463)"
    .range("V16:V194").Formula = "=IFERROR(ABS(P16-J16)/P16,NA())"
    
    .range("E5").Value = Format(month1, "mmmm yyyy")
    .range("E6").Value = Format(Month2, "mmmm yyyy")
    .range("E7").Value = Format(Month3, "mmmm yyyy")
    .range("E8").Value = mQtr & "-" & mYear
End With

'clear Field Points Data
With wb.Worksheets("Field Points Data")
    .range("A9:E196").ClearContents
    .range("G9:K46").ClearContents
End With

'correct an anormaly in tab name (two spaces)
If TabExists("QAQC  Notes", wb) = True Then
    Sheets("QAQC  Notes").Name = "QAQC Notes"
End If

'update the month-year in table
With wb.Sheets("QAQC Notes")
    .range("A7").Value = Format(month1, "mmmm yyyy")
    .range("A8").Value = Format(Month2, "mmmm yyyy")
    .range("A9").Value = Format(Month3, "mmmm yyyy")
    .range("A13:E32").ClearContents
End With

'update datetime in "Flow Data"
With wb.Worksheets("Flow Data")
    'clear contents of previous dates
    .range("A" & Row1, "Y" & .Cells(.Rows.count, "A").End(xlUp).Row).ClearContents
    'fill range (fastest way)
    'for comparision of other methods, refer to [populate_time_range] test()
    .range("A" & Row1).Value = month1
    .range("A" & Row1 + 1, "A" & Row4).Formula = "=A" & Row1 & "+(" & intvl & "/60)/24"
    .range("A" & Row1, "A" & Row4).NumberFormat = "mm/dd/yyyy hh:mm:ss"
    .range("I" & Row1, "I" & Row4).Formula = "=IF(B14>0,If(K14<>" & Chr(34) & Chr(34) & ",VLOOKUP(B14,'Area vs. depth table'!$A:$C,3,TRUE)-VLOOKUP(K14,'Area vs. depth table'!$A:$C,3,TRUE),IF(B14>0,VLOOKUP(B14,'Area vs. depth table'!$A:$C,3,TRUE)))," & Chr(34) & Chr(34) & ")"
    .range("J" & Row1, "J" & Row4).Formula = "=IF(AND(I14>0,C14>0),I14*C14*0.64632," & Chr(34) & Chr(34) & ")"
    .range("U" & Row1, "U" & Row4).Formula = "=IF(AND(L14=" & Chr(34) & Chr(34) & ", P14=" & Chr(34) & Chr(34) & "), IF(B14>0,B14," & Chr(34) & Chr(34) & "), (IF(L14=P14, L14, P14)))"
    .range("V" & Row1, "V" & Row4).Formula = "=IF(AND(M14=" & Chr(34) & Chr(34) & ", Q14=" & Chr(34) & Chr(34) & "), IF(C14>0,C14," & Chr(34) & Chr(34) & "), (IF(M14=Q14, M14, Q14)))"
    .range("W" & Row1, "W" & Row4).Formula = "=IF(AND(N14=" & Chr(34) & Chr(34) & ", R14=" & Chr(34) & Chr(34) & "), IF(J14>0,J14," & Chr(34) & Chr(34) & "), (IF(N14=R14, N14, R14)))"
    .range("X" & Row1, "X" & Row4).Formula = "=IF(AND(F14<>" & Chr(34) & Chr(34) & ",F14>0),F14," & Chr(34) & Chr(34) & ")"
    .range("Y" & Row1, "Y" & Row4).Formula = "=IF(OR(B14=" & Chr(34) & Chr(34) & ",C14=" & Chr(34) & Chr(34) & ",B14<=0,C14<=0),$AA$3,IF(W14=" & Chr(34) & Chr(34) & ",$AA$2,IF(AND(U14=D14,B14<>D14),$AA$5,IF(OR(L14<>" & Chr(34) & Chr(34) & ",M14<>" & Chr(34) & Chr(34) & ",N14<>" & Chr(34) & Chr(34) & "),$AA$9,IF(OR(P14<>" & Chr(34) & Chr(34) & ",Q14<>" & Chr(34) & Chr(34) & ",R14<>" & Chr(34) & Chr(34) & "),$AA$10,IF(K14<>" & Chr(34) & Chr(34) & ",$AA$8,$AA$4))))))"
    If intvl = 15 Then
    .range("AA" & Row1, "AA" & Row4).Formula = "='Rainfall Data'!B2"
    End If
    
'update the month-year in the Percent Recovery in Flow Data
    .range("H5").Value = Format(startDate, "mmmm")
    .range("H6").Value = Format(DateAdd("m", 1, startDate), "mmmm")
    .range("H7").Value = Format(DateAdd("m", 2, startDate), "mmmm")
'change the Percent Recovery in Flow Data
    .range("I5").Value = "=(COUNT(U" & Row1 & ":U" & (Row2 - 1) & "))/(COUNT(A" & Row1 & ":A" & (Row2 - 1) & "))"
    .range("I6").Value = "=(COUNT(U" & Row2 & ":U" & (Row3 - 1) & "))/(COUNT(A" & Row2 & ":A" & (Row3 - 1) & "))"
    .range("I7").Value = "=(COUNT(U" & Row3 & ":U" & (Row4 - 1) & "))/(COUNT(A" & Row3 & ":A" & (Row4 - 1) & "))"
    .range("j5").Value = "=(COUNT(w" & Row1 & ":w" & (Row2 - 1) & "))/(COUNT(A" & Row1 & ":A" & (Row2 - 1) & "))"
    .range("j6").Value = "=(COUNT(w" & Row2 & ":w" & (Row3 - 1) & "))/(COUNT(A" & Row2 & ":A" & (Row3 - 1) & "))"
    .range("j7").Value = "=(COUNT(w" & Row3 & ":w" & (Row4 - 1) & "))/(COUNT(A" & Row3 & ":A" & (Row4 - 1) & "))"
    .range("k5").Value = "=(COUNT(v" & Row1 & ":v" & (Row2 - 1) & "))/(COUNT(A" & Row1 & ":A" & (Row2 - 1) & "))"
    .range("k6").Value = "=(COUNT(v" & Row2 & ":v" & (Row3 - 1) & "))/(COUNT(A" & Row2 & ":A" & (Row3 - 1) & "))"
    .range("k7").Value = "=(COUNT(v" & Row3 & ":v" & (Row4 - 1) & "))/(COUNT(A" & Row3 & ":A" & (Row4 - 1) & "))"
End With

'clear contents of previous data
With wb.Worksheets("Area vs. depth table")
    .range("A3:C" & .Cells(.Rows.count, 1).End(xlUp).Row).ClearContents
End With

'clear contents of previous data
With wb.Worksheets("Rainfall Data")
    .range("A:B").ClearContents
    .Columns(1).NumberFormat = "m/d/yyyy h:m"
End With


'Part II: Charts

'change Charts names
With wb
    For i = 1 To 18
    .Charts(i).Name = "chart" & i
    Next i
    j = 1
    For i = 0 To 2
        .Charts(j).Name = MonthName(startMonth + i, True) & " SP (Flow)"
        .Charts(j + 1).Name = MonthName(startMonth + i, True) & " SP CORR (Flow)"
        .Charts(j + 2).Name = MonthName(startMonth + i, True) & " SP (Vel)"
        .Charts(j + 3).Name = MonthName(startMonth + i, True) & " SP CORR (Vel)"
        .Charts(j + 4).Name = MonthName(startMonth + i, True) & " TS"
        .Charts(j + 5).Name = MonthName(startMonth + i, True) & " TS CORR"
        j = j + 6
    Next i
End With

' change the monthly charts' data range and scale

Call QAQC_SP(month1, Month2, Row1, (Row2 - 1))
Call QAQC_TS(month1, Month2, Row1, (Row2 - 1), intvl)

Call QAQC_SP(Month2, Month3, Row2, (Row3 - 1))
Call QAQC_TS(Month2, Month3, Row2, (Row3 - 1), intvl)

Call QAQC_SP(Month3, Month4, Row3, (Row4 - 1))
Call QAQC_TS(Month3, Month4, Row3, (Row4 - 1), intvl)

' change the ALL charts data range and scale
Call QAQC_SP("ALL", Month4, Row1, (Row4 - 1))
Call QAQC_TS("ALL", Month4, Row1, (Row4 - 1), intvl)
'change the rest 4 charts' data range only, scales are not changed
Call QAQC_misc(Row1, (Row4 - 1))

'save the generated new template
wb.Save
'unload the wb object
Set wb = Nothing
Application.ScreenUpdating = True

MsgBox fso.GetFileName(tgtPath) & " has been created.", vbInformation, "Success!"
Call loginfo(tgtName, fso.GetFileName(tgtPath) & " created")
End Sub

Private Sub QAQC_misc(startRow, endRow)
'Hao Zhang @ 2015.1.31
'separate from QAtempGenBtn_click()
'change data range for the charts: SP Flow Vs Level 1&2, SP Velocity Vs Level 1&2, SP Velocity Vs Level 1&2, SP Raw Flow Vs Corr Flow
'scales do not need to be changed
With Charts("SP Flow Vs Level 1&2")
'level 1
    .SeriesCollection("Primary Level").XValues = "='Flow Data'!$B$" & startRow & ":$B$" & endRow
'Flow
    .SeriesCollection("Primary Level").Values = "='Flow Data'!$W$" & startRow & ":$W$" & endRow
'level 2
    .SeriesCollection("Redundant Level").XValues = "='Flow Data'!$D$" & startRow & ":$D$" & endRow
'Flow
    .SeriesCollection("Redundant Level").Values = "='Flow Data'!$W$" & startRow & ":$W$" & endRow
End With

With Charts("SP Velocity Vs Level 1&2")
'level 1
    .SeriesCollection("Primary Level").XValues = "='Flow Data'!$B$" & startRow & ":$B$" & endRow
'Vel 1
    .SeriesCollection("Primary Level").Values = "='Flow Data'!$V$" & startRow & ":$V$" & endRow
'level 1
    .SeriesCollection("Redundant Level").XValues = "='Flow Data'!$D$" & startRow & ":$D$" & endRow
'Vel 2
    .SeriesCollection("Redundant Level").Values = "='Flow Data'!$V$" & startRow & ":$V$" & endRow
End With

With Charts("SP Raw Flow Vs Corr Flow")
'RAW Flow
    .SeriesCollection("Calc Flow Vs Raw FLow").XValues = "='Flow Data'!$G$" & startRow & ":$G$" & endRow
'Cal Flow
    .SeriesCollection("Calc Flow Vs Raw FLow").Values = "='Flow Data'!$J$" & startRow & ":$J$" & endRow
End With
End Sub
Private Sub QAQC_SP(startDate, endDate, startRow, endRow)
'Hao Zhang @ 2015.1.31
'updates the data range and time scale for (monthly and ALL) SP charts

'add a special case for ALL SP
If startDate = "ALL" Then
    tabName = "ALL"
    startDate = DateAdd("m", -3, endDate)
Else
    tabName = MonthName(Month(startDate), True)
End If

With Charts(tabName & " SP (Flow)")
'level
    .SeriesCollection("Monitored Data").XValues = "='Flow Data'!$B$" & startRow & ":$B$" & endRow
'Flow
    .SeriesCollection("Monitored Data").Values = "='Flow Data'!$G$" & startRow & ":$G$" & endRow
End With

With Charts(tabName & " SP CORR (Flow)")
'level CORR
    .SeriesCollection("Monitored Data").XValues = "='Flow Data'!$U$" & startRow & ":$U$" & endRow
'Flow CORR
    .SeriesCollection("Monitored Data").Values = "='Flow Data'!$W$" & startRow & ":$W$" & endRow
End With

'add a special case for ALL SP
If tabName <> "ALL" Then
    With Charts(tabName & " SP (Vel)")
    'level
        .SeriesCollection("Monitored Data").XValues = "='Flow Data'!$B$" & startRow & ":$B$" & endRow
    'Vel
        .SeriesCollection("Monitored Data").Values = "='Flow Data'!$C$" & startRow & ":$C$" & endRow
    End With
    
    With Charts(tabName & " SP CORR (Vel)")
    'level CORR
        .SeriesCollection("Monitored Data").XValues = "='Flow Data'!$U$" & startRow & ":$U$" & endRow
    'Vel CORR
        .SeriesCollection("Monitored Data").Values = "='Flow Data'!$V$" & startRow & ":$V$" & endRow
    End With
End If

End Sub

Private Sub QAQC_TS(startDate, endDate, startRow, endRow, intvl)
'Hao Zhang @ 2015.1.31
'updates the data range and time scale for (monthly and ALL) TS charts

'add a special case for ALL TS
If startDate = "ALL" Then
    tabName = "ALL"
    startDate = DateAdd("m", -3, endDate)
Else
    tabName = MonthName(Month(startDate), True)
End If

'monthly TS
With Charts(tabName & " TS")
    'hyetograph
    With .ChartObjects("Rain").Chart
        If intvl = 15 Then
            'adjust time range
            .SeriesCollection("Rainfall").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
            .SeriesCollection("Rainfall").Values = "='Flow Data'!$AA$" & startRow & ":$AA$" & endRow
            'adjust time scale
            .Axes(xlCategory).MaximumScale = endDate
            .Axes(xlCategory).MinimumScale = startDate
        Else
            'for 2min and 5min data, use the rainfall sheet instead
            'adjust time range
            .SeriesCollection("Rainfall").XValues = "='Rainfall Data'!$A$2:$A$" & endRow - startRow + 1
            .SeriesCollection("Rainfall").Values = "='Rainfall Data'!$B$2:$B$" & endRow - startRow + 1
            'adjust time scale
            .Axes(xlCategory).MaximumScale = endDate
            .Axes(xlCategory).MinimumScale = startDate
        End If
    End With
'hydrograph
'Dtime
    .SeriesCollection("Level 2").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Level2
    .SeriesCollection("Level 2").Values = "='Flow Data'!$D$" & startRow & ":$D$" & endRow
'Dtime
    .SeriesCollection("Level 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Level1
    .SeriesCollection("Level 1").Values = "='Flow Data'!$B$" & startRow & ":$B$" & endRow
'Dtime
    .SeriesCollection("Vel 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Vel1
    .SeriesCollection("Vel 1").Values = "='Flow Data'!$C$" & startRow & ":$C$" & endRow
'Dtime
    .SeriesCollection("Vel 2").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Vel2
    .SeriesCollection("Vel 2").Values = "='Flow Data'!$E$" & startRow & ":$E$" & endRow
'Dtime
    .SeriesCollection("Flow 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Flow1
    .SeriesCollection("Flow 1").Values = "='Flow Data'!$G$" & startRow & ":$G$" & endRow
'Dtime
    .SeriesCollection("Flow 2").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
'Flow2
    .SeriesCollection("Flow 2").Values = "='Flow Data'!$H$" & startRow & ":$H$" & endRow
'adjust time scale
    .Axes(xlCategory).MaximumScale = endDate
    .Axes(xlCategory).MinimumScale = startDate
End With


With Charts(tabName & " TS CORR")
    'hyetograph for ALL TS CORR only
    If tabName = "ALL" Then
        With .ChartObjects("Rain").Chart
            If intvl = 15 Then
                'for 15min data
                'adjust time range
                .SeriesCollection("Rainfall").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
                .SeriesCollection("Rainfall").Values = "='Flow Data'!$AA$" & startRow & ":$AA$" & endRow
                'adjust time scale
                .Axes(xlCategory).MaximumScale = endDate
                .Axes(xlCategory).MinimumScale = startDate
            Else
                'for 2min and 5min data, use the rainfall sheet instead
                'adjust time range
                .SeriesCollection("Rainfall").XValues = "='Rainfall Data'!$A$2:$A$" & endRow - startRow + 1
                .SeriesCollection("Rainfall").Values = "='Rainfall Data'!$B$2:$B$" & endRow - startRow + 1
                'adjust time scale
                .Axes(xlCategory).MaximumScale = endDate
                .Axes(xlCategory).MinimumScale = startDate
            End If
        End With
        'hydrograph
        'Dtime
        .SeriesCollection("Corrected Level (in)").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
        'Level1
        .SeriesCollection("Corrected Level (in)").Values = "='Flow Data'!$U$" & startRow & ":$U$" & endRow
        'Dtime
        .SeriesCollection("Corrected Flow (mgd)").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
        'Flow1
        .SeriesCollection("Corrected Flow (mgd)").Values = "='Flow Data'!$W$" & startRow & ":$W$" & endRow
        'adjust time scale
    Else
        'hydrograph
        'Dtime
        .SeriesCollection("Level 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
        'Level1
        .SeriesCollection("Level 1").Values = "='Flow Data'!$U$" & startRow & ":$U$" & endRow
        'Dtime
        .SeriesCollection("Flow 1").XValues = "='Flow Data'!$A$" & startRow & ":$A$" & endRow
        'Flow1
        .SeriesCollection("Flow 1").Values = "='Flow Data'!$W$" & startRow & ":$W$" & endRow
        'adjust time scale
        .Axes(xlCategory).MaximumScale = endDate
        .Axes(xlCategory).MinimumScale = startDate
    End If
End With

End Sub



'*******************************************************************
'***********************QAQC elaboration****************************
'*******************************************************************


Private Sub QASiteInfoBtn_Click()
'Hao Zhang @ 2015.2.6
'create a site QA sheet if not exist, fill "site info", "historical FP data" and "Area vs. depth data" from previous quarterly sheet

'Hao Zhang @ 2015.2.4
'copy and paste headings from previous quarter QA sheet
'affecting tab: Site Info

Dim srcWb As Workbook
Dim tgtWb As Workbook

'Hao Zhang @ 2015.2.9

''get the selected item's name and row number
'For iSite = 0 To chosenSitesLB.ListCount - 1
'    If chosenSitesLB.Selected(iSite) = True Then
'        siteName = chosenSitesLB.List(iSite, 1)
'        Row = chosenSitesLB.List(iSite, 0)
'        'only the first selected item is returned
'        Exit For
'    End If
'Next
'
'QtrYr = QAqtrCB.Value & "-" & Right(QAyrCB.Value, 2)
'Col = MyWs.Rows(1).Find(QtrYr).Column
'QAsheet = GetQApath(siteName, QtrYr)
''if QA sheet doesn't exist, then create a new QA sheet from a template, and write the path in CurSitesTbl
'If QAsheet = "" Then
'    QAsheet = MyWs.Cells(Row, 6).Value & "\QAQC\" & siteName & " (" & QtrYr & ").xlsm"
'    tempQAsheet = "C:\Users\hao.zhang\Desktop\QAQC templates\Template (Q1-15)_HZ20150115.xlsm"
'    FileCopy tempQAsheet, QAsheet
'    MyWs.Cells(Row, Col).Value = QAsheet
'End If

Call loadSite


If MyWs.Cells(siteRow, QtrYr_Col - 1).Value <> "" Then
    If IsWorkBookOpen(MyWs.Cells(siteRow, QtrYr_Col - 1).Value) = False Then
        Set srcWb = Workbooks.Open(fileName:=MyWs.Cells(siteRow, QtrYr_Col - 1).Value, ReadOnly:=True)
    Else
        Set srcWb = Workbooks(fso.GetFileName(MyWs.Cells(siteRow, QtrYr_Col - 1).Value))
    End If
    'hide the src workbook
    srcWb.Windows(1).visible = False
'quit the sub when no previous QA sheet are available
Else
    MsgBox " this must be a new site, please add site info manually."
    Exit Sub
End If

Set tgtWb = QAwb

'With srcWb.Worksheets("Site Info")
'    Dim EOR As Integer
'    EOR = .Cells(.Rows.Count, "C").End(xlUp).Row
'    'get row/col number
'    With .Range("A1:D16")
'        sLM_row1 = .Find("Last Modified:").Row
'        sLM_row2 = LM_row1 + 1
'        sSN_row = .Find("Site Name:").Row
'        sDL_row = .Find("Description of Location:").Row
'        sTM_row = .Find("Type of Monitor:").Row
'        sPL_row = .Find("Primary Level Sensor:").Row
'        sRL_row = .Find("Redundant Level Sensor:").Row
'        sPT_row = .Find("Pipe Type:").Row
'        sPD_row = .Find("Pipe Depth/Dimensions:").Row
'        sRG_row = .Find("RG#:").Row
'        sDA_row = .Find("Drainage Area (acre):").Row
'        sRF_row = .Find("Related Files:").Row
'        sCol1 = .Find("Last Modified:").Column + 1
'        sCol2 = Col1 + 1
'    End With
'End With

Dim HeadingRng As range
Dim FPRng As range

With srcWb.Worksheets("Site Info")
    Set HeadingRng = .range("A2:A20")
    Set FPRng = .range("A15:Z25")
End With

With tgtWb.Worksheets("Site Info")
    'copy site info
    'If .Range("B2:B13") = "" Then
        .range("B2:C3").Value = HeadingRng.Find("Last Modified:", SearchDirection:=xlNext).Offset(0, 1).Resize(2, 2).Value
        For iRow = 4 To 13
            searchTerm = .Cells(iRow, 1).Value
            .range("B" & iRow, "C" & iRow).Value = HeadingRng.Find(searchTerm).Offset(0, 1).Resize(1, 2).Value
        DoEvents
        Next
        .range("G2").Value = HeadingRng.Find("Date Installed: ").Offset(0, 1).Value
        .range("G3").Value = HeadingRng.Find("Date Removed: ").Offset(0, 1).Value
        tgtStartRow = .range("15:25").Find("Date Time").Row + 1
    'End If
    
    'copy field points
    'If .Range("C16:C21") = "" Then
    
    '''[debug]: silt vs silt/ Debris
        For jCol = 3 To 21
            If .Cells(tgtStartRow, jCol).HasFormula = False Then
                searchTerm = .Cells(tgtStartRow - 1, jCol).Value
                srcCol = FPRng.Find(searchTerm, lookat:=xlWhole).Column
                srcStartRow = FPRng.Find(searchTerm, lookat:=xlWhole).Row + 1
                srcEndRow = srcWb.Worksheets("Site Info").Cells(.Rows.count, jCol).End(xlUp).Row
                .range(.Cells(tgtStartRow, jCol).Address, .Cells(srcEndRow - srcStartRow + tgtStartRow, jCol).Address).Value = _
                srcWb.Worksheets("Site Info").range(.Cells(srcStartRow, srcCol).Address, .Cells(srcEndRow, srcCol).Address).Value
            End If
        DoEvents
        Next
        'fix number format for the two columns
        .range("F16:G16").Resize(srcEndRow - srcStartRow, 2).NumberFormat = "General"
    'End If
End With
    
    'get row/col number
'    With .Range("A1:D16")
'        LM_row1 = .Find("Last Modified:").Row
'        LM_row2 = LM_row1 + 1
'        SN_row = .Find("Site Name:").Row
'        DL_row = .Find("Description of Location:").Row
'        TM_row = .Find("Type of Monitor:").Row
'        PL_row = .Find("Primary Level Sensor:").Row
'        RL_row = .Find("Redundant Level Sensor:").Row
'        PT_row = .Find("Pipe Type:").Row
'        PD_row = .Find("Pipe Depth/Dimensions:").Row
'        RG_row = .Find("RG#:").Row
'        DA_row = .Find("Drainage Area (acre):").Row
'        RF_row = .Find("Related Files:").Row
'        Col1 = .Find("Last Modified:").Column + 1
'        Col2 = Col1 + 1
'    End With
'
'    .Cells(LM_row1, Col1).Value = srcWb.Worksheets("Site Info").Cells(sLM_row1, sCol1).Value
'    .Cells(LM_row2, Col1).Value = srcWb.Worksheets("Site Info").Cells(sLM_row2, sCol1).Value
'    .Cells(LM_row1, Col2).Value = srcWb.Worksheets("Site Info").Cells(sLM_row1, sCol2).Value
'    .Cells(LM_row2, Col2).Value = srcWb.Worksheets("Site Info").Cells(sLM_row2, sCol2).Value
'    .Cells(SN_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sSN_row, sCol1).Value
'    .Cells(DL_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sDL_row, sCol1).Value
'    .Cells(TM_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sTM_row, sCol1).Value
'    .Cells(PL_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sPL_row, sCol1).Value
'    .Cells(RL_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sRL_row, sCol1).Value
'    .Cells(PT_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sPT_row, sCol1).Value
'    .Cells(PD_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sPD_row, sCol1).Value
'    .Cells(RG_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sRG_row, sCol1).Value
'    .Cells(DA_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sDA_row, sCol1).Value
'    .Cells(RF_row, Col1).Value = srcWb.Worksheets("Site Info").Cells(sRF_row, sCol1).Value
'
'    'copy headings
'    .Range("B2:C3").Value = srcWb.Worksheets("Site Info").Range("B2:C3").Value
'    .Range("B4:C5").Value = srcWb.Worksheets("Site Info").Range("B5:C6").Value
'    .Range("B6:C12").Value = srcWb.Worksheets("Site Info").Range("B8:C14").Value
'    .Range("B13:C13").Value = srcWb.Worksheets("Site Info").Range("B16:C16").Value
'    'copy existing FP
'    .Range("C16:H" & (EOR - 23 + 16)).Value = srcWb.Worksheets("Site Info").Range("C23:H" & EOR).Value
'    .Range("K16:N" & (EOR - 23 + 16)).Value = srcWb.Worksheets("Site Info").Range("K23:N" & EOR).Value
'    .Range("Q16:U" & (EOR - 23 + 16)).Value = srcWb.Worksheets("Site Info").Range("Q23:U" & EOR).Value
'End With

'copying the pipe profile
''in future, this may be pulled directly from pipe profile file

With tgtWb.Worksheets("Area vs. depth table")
    'If .Range("A:C").Value <> srcWb.Worksheets("Area vs. depth table").Range("A:C").Value Then
        .range("A:C").Value = srcWb.Worksheets("Area vs. depth table").range("A:C").Value
    'End If
End With

srcWb.Close savechanges:=False
Call loginfo(siteName, fso.GetFileName(QAsheet) & " created/imported")


End Sub

Private Sub QAappFlowBtn_Click()
'HaoZhang @ 2015.2.5
'append one folder load of Raw data into Flow Data in QA sheet
'append one folder load of field point data into Site Info in QA sheet


'Stop the screen from flashing
Application.ScreenUpdating = False

Call loadSite
'For iSite = 0 To chosenSitesLB.ListCount - 1
'    If chosenSitesLB.Selected(iSite) = True Then
'        siteName = chosenSitesLB.List(iSite, 1)
'        Row = chosenSitesLB.List(iSite, 0)
'        Exit For
'    End If
'Next
'
'QtrYr = QAqtrCB.Value & "-" & Right(QAyrCB.Value, 2)
'Col = MyWs.Rows(1).Find(QtrYr).Column
'QAsheet = GetQApath(siteName, QtrYr)
'
''open the site QA sheet
'If fso.FileExists(QAsheet) = True Then
'    Application.Workbooks.Open fileName:=QAsheet
'Else
'    a = MsgBox("Sorry, can't find the QA file, please specify:", vbOKCancel, "Warning")
'    If a = vbOK Then
'        QAsheet = GetFile(rootPath)
'        Application.Workbooks.Open fileName:=QAsheet
'    Else
'        Exit Sub
'    End If
'End If

Dim sitePath As String
Dim rawFldr As String
Dim rawPath As String

sitePath = MyWs.Cells(siteRow, siteFldr_Col).Value
rawFldr = RawDateTB.Value

'dealing with naming variations
If fso.FolderExists(sitePath & "\Raw Data\") = True Then
    rawPath = sitePath & "\Raw Data\" & rawFldr
Else
    rawPath = sitePath & "\RawData\" & rawFldr
End If

If fso.FolderExists(rawPath) = False Then
    a = MsgBox("Cannot find the Raw Data folder, please spicify:", vbOKCancel, "Warning")
    If a = vbOK Then
        rawPath = GetFldr(sitePath)
    Else
        Exit Sub
    End If
End If

'find 15min data
If Not siteName Like "*min*" Then
    'find the three raw data files
    For Each fl In fso.GetFolder(rawPath).Files
        'if this doesn't work, use the instr() instead
        If fl.Name Like "*Excel*" And Not fl.Name Like "*Minute*" Then
            If Not fl.Name Like "*Redundant*" Then
                file_LVF = fl.path
            Else
                file_Redun = fl.path
            End If
        Else
            If fl.Name Like "*Electronic Fieldbook*" Then
                file_FP = fl.path
            End If
        End If
    Next
'find 2min or 5min data
Else
    'find the three files based on keywords
    For Each fl In fso.GetFolder(rawPath).Files
        If fl.Name Like "*Excel*" And fl.Name Like "*Minute*" Then
            If Not fl.Name Like "*Redundant*" Then
                file_LVF = fl.path
            Else
                file_Redun = fl.path
            End If
        Else
            If fl.Name Like "*Electronic Fieldbook*" Then
                file_FP = fl.path
            End If
        End If
    Next
End If


'check if the src files exists
If fso.FileExists(file_LVF) = False Then
    a = MsgBox("Cannot find the Main file, please spicify:", vbOKCancel, "Warning")
    If a = vbOK Then
        Set LVFwb = Workbooks.Open(fileName:=GetFile(rawPath))
    Else
        MsgBox "Operation canceled"
        Exit Sub
    End If
Else
    Set LVFwb = Workbooks.Open(fileName:=file_LVF)
End If

If fso.FileExists(file_Redun) = False Then
    a = MsgBox("Cannot find the Redundant file, please spicify:", vbOKCancel, "Warning")
    If a = vbOK Then
        Set Redwb = Workbooks.Open(fileName:=GetFile(rawPath))
    Else
        MsgBox "Operation canceled"
        Exit Sub
    End If
Else
    Set Redwb = Workbooks.Open(fileName:=file_Redun)
End If

If fso.FileExists(file_FP) = False Then
    a = MsgBox("Cannot find the Fieldbook file, please spicify:", vbOKCancel, "Warning")
    If a = vbOK Then
        Set FPwb = Workbooks.Open(fileName:=GetFile(rawPath))
    Else
        MsgBox "Operation is canceled."
        Exit Sub
    End If
Else
    Set FPwb = Workbooks.Open(fileName:=file_FP)
End If

'If fso.FileExists(file_Redun) = False Then
'    MsgBox "can't find the redundant data file"
'    Exit Sub
'End If
'If fso.FileExists(file_FP) = False Then
'    MsgBox "can't find the field point data file"
'    Exit Sub
'End If
'
''identify the QA sheet
'If QAsheet <> "" Then
'    If IsWorkBookOpen(QAsheet) = True Then
'        Set QAwb = Workbooks(fso.GetFileName(QAsheet))
'    Else
'        Set QAwb = Workbooks.Open(fileName:=QAsheet)
'    End If
'End If
'
'Set LVFwb = Workbooks.Open(fileName:=file_LVF)
'Set Redwb = Workbooks.Open(fileName:=file_Redun)
'Set FPwb = Workbooks.Open(fileName:=file_FP)
'
'find the start/end row/Date in Raw data files
'assuming LVF and Redundant files are synced

With LVFwb.Sheets(1)
    RawStartRow = 2
    RawStartDate = .Cells(RawStartRow, 1).Value
    RawEndRow = .Cells(Rows.count, 1).End(xlUp).Row - 4  ''take out the summary rows on bottom
    RawEndDate = .Cells(RawEndRow, 1).Value
End With

'find the start/end row/Date in Raw data files
'assuming LVF and Redundant files are synced
With FPwb.Sheets(1)
    LastDate = QAwb.Worksheets("Site Info").Cells(.Rows.count, 3).End(xlUp).Offset(0, -1).Value
    lastRow = QAwb.Worksheets("Site Info").Cells(.Rows.count, 3).End(xlUp).Offset(0, -1).Row
    FPStartRow = .Columns(1).Find(Format(LastDate, "m/d/yyyy h:mm"), LookIn:=xlValues, lookat:=xlWhole).Row + 1
    FPStartDate = .Cells(FPStartRow, 1).Value
    FPEndRow = .Cells(.Rows.count, 1).End(xlUp).Row
    FPEndDate = .Cells(FPEndRow, 1).Value
    QAFPStartRow = lastRow + 1
    QAFPEndRow = QAFPStartRow + FPEndRow - FPStartRow
    With .Rows("1:3")
        rDT_Col = .Find("Date").Column
        rTM_Col = .Find("Time").Column
        rMT_Col = .Find("Meter Time").Column
        
        'these two columns have problems
        If Not .Find("Inside Crown to Water Surface (in)") Is Nothing Then
            rCW_Col = .Find("Inside Crown to Water Surface (in)").Column
        End If
        If Not .Find("Water Surface to Silt (in)") Is Nothing Then
            rWS_Col = .Find("Water Surface to Silt (in)").Column
        End If
        
        rFD_Col = .Find("Field Depth").Column
        rFV_Col = .Find("Field Vel.").Column
        rUD_Col = .Find("UpDepth (inches)").Column
        rPD_Col = .Find("Pdepth (inches)").Column
        'for some sites, the Sdepth is called Pdepth 2, use this method to bypass the inconsistency
        rSD_Col = rPD_Col + 1
        rMV_Col = .Find("Meter Vel.").Column
        rST_Col = .Find("Silt").Column
        rBT_Col = .Find("Battery").Column
        rIN_Col = .Find("Initials").Column
        rSC_Col = .Find(" Service Comments").Column
    End With
End With


With QAwb.Worksheets("Flow Data")
    'find the corresponding start/end row/date in QA sheet, and the column numbers of each parameter
    QAStartDate = .Columns(1).Find(Format(RawStartDate, "mm/dd/yyyy hh:mm:ss"), LookIn:=xlValues, lookat:=xlWhole).Value
    QAStartRow = .Columns(1).Find(Format(RawStartDate, "mm/dd/yyyy hh:mm:ss"), LookIn:=xlValues, lookat:=xlWhole).Row
    QAEndDate = .Columns(1).Find(Format(RawEndDate, "mm/dd/yyyy hh:mm:ss"), LookIn:=xlValues, lookat:=xlWhole).Value
    QAEndRow = .Columns(1).Find(Format(RawEndDate, "mm/dd/yyyy hh:mm:ss"), LookIn:=xlValues, lookat:=xlWhole).Row
    'find the first empty row in QA sheet
    EOR = .Cells(.Rows.count, "B").End(xlUp).Row + 1
    L1_Col = .Rows("10:16").Find("Level 1").Column
    V1_Col = .Rows("10:16").Find("Vel 1").Column
    L2_Col = .Rows("10:16").Find("Level 2").Column
    TP_Col = .Rows("10:16").Find("Temp").Column
    F1_Col = .Rows("10:16").Find("Flow 1").Column
    ST_Col = .Rows("10:16").Find("Silt / Debris").Column
    
    'determine if Raw data need to be appended
    'cond1: appending data is not connected to existing data
    If EOR < QAStartRow Then
        a = MsgBox("there is a gap between current QA sheet and the appending data." & Chr(13) & "Continue?", vbOKCancel)
        If a = vbCancel Then
            Exit Sub
        End If
        'EOR = .Columns(1).Find(RawStartDate).Row
    Else
    'cond2: appending data is overlapping with existing data
        If EOR > QAStartRow Then
            If EOR = QAEndRow + 1 Then
                a = MsgBox("Raw data already exists, overwrite?", vbYesNo, "Warning")
                If a = vbNo Then
                    Exit Sub
                End If
            Else
                'partially import new data
                RawStartDate = .Cells(EOR, "A").Value
                RawStartRow = LVFwb.Columns(1).Find(RawStartDate).Row
                QAStartDate = .Cells(EOR, "A").Value
                QAStartRow = EOR
            End If
        End If
    End If
    
    'copy data from Raw to QA
    Dim copyRange As range
    Set copyRange = LVFwb.Sheets(1).range(Cells(RawStartRow, 2).Address, Cells(RawEndRow, 2).Address)
    copyRange.Copy Destination:=.range(Cells(QAStartRow, L1_Col).Address, Cells(QAEndRow, L1_Col).Address)
    Set copyRange = LVFwb.Sheets(1).range(Cells(RawStartRow, 3).Address, Cells(RawEndRow, 3).Address)
    copyRange.Copy Destination:=.range(Cells(QAStartRow, V1_Col).Address, Cells(QAEndRow, V1_Col).Address)
    Set copyRange = LVFwb.Sheets(1).range(Cells(RawStartRow, 4).Address, Cells(RawEndRow, 4).Address)
    copyRange.Copy Destination:=.range(.Cells(QAStartRow, F1_Col).Address, .Cells(QAEndRow, F1_Col).Address)
    Set copyRange = Redwb.Sheets(1).range(Cells(RawStartRow, 2).Address, Cells(RawEndRow, 2).Address)
    copyRange.Copy Destination:=.range(.Cells(QAStartRow, L2_Col).Address, .Cells(QAEndRow, L2_Col).Address)
    Set copyRange = LVFwb.Sheets(1).range(Cells(RawStartRow, 5).Address, Cells(RawEndRow, 5).Address)
    copyRange.Copy Destination:=.range(.Cells(QAStartRow, TP_Col).Address, .Cells(QAEndRow, TP_Col).Address)
    Set copyRange = LVFwb.Sheets(1).range(Cells(RawStartRow, 6).Address, Cells(RawEndRow, 6).Address)
    copyRange.Copy Destination:=.range(.Cells(QAStartRow, ST_Col).Address, .Cells(QAEndRow, ST_Col).Address)
End With


'copy data from FP to QA
With QAwb.Worksheets("Site Info")
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rDT_Col).Address, Cells(FPEndRow, rDT_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 3).Address, Cells(QAFPEndRow, 3).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rTM_Col).Address, Cells(FPEndRow, rTM_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 4).Address, Cells(QAFPEndRow, 4).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rMT_Col).Address, Cells(FPEndRow, rMT_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 5).Address, Cells(QAFPEndRow, 5).Address)
    If rCW_Col <> "" Then
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rCW_Col).Address, Cells(FPEndRow, rCW_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 6).Address, Cells(QAFPEndRow, 6).Address)
    End If
    If rWS_Col <> "" Then
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rWS_Col).Address, Cells(FPEndRow, rWS_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 7).Address, Cells(QAFPEndRow, 7).Address)
    End If
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rFD_Col).Address, Cells(FPEndRow, rFD_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 8).Address, Cells(QAFPEndRow, 8).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rFV_Col).Address, Cells(FPEndRow, rFV_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 11).Address, Cells(QAFPEndRow, 11).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rUD_Col).Address, Cells(FPEndRow, rUD_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 12).Address, Cells(QAFPEndRow, 12).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rPD_Col).Address, Cells(FPEndRow, rPD_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 13).Address, Cells(QAFPEndRow, 13).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rSD_Col).Address, Cells(FPEndRow, rSD_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 14).Address, Cells(QAFPEndRow, 14).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rMV_Col).Address, Cells(FPEndRow, rMV_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 17).Address, Cells(QAFPEndRow, 17).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rST_Col).Address, Cells(FPEndRow, rST_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 18).Address, Cells(QAFPEndRow, 18).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rBT_Col).Address, Cells(FPEndRow, rBT_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 19).Address, Cells(QAFPEndRow, 19).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rIN_Col).Address, Cells(FPEndRow, rIN_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 20).Address, Cells(QAFPEndRow, 20).Address)
    Set copyRange = FPwb.Sheets(1).range(Cells(FPStartRow, rSC_Col).Address, Cells(FPEndRow, rSC_Col).Address)
    copyRange.Copy Destination:=.range(Cells(QAFPStartRow, 21).Address, Cells(QAFPEndRow, 21).Address)
'    Set copyRange = FPwb.Sheets(1).Range(Cells(FPStartRow, 5).Address, Cells(FPEndRow, 5).Address)
'    copyRange.Copy Destination:=.Range(Cells(QAFPStartRow, 8).Address, Cells(QAFPEndRow, 8).Address)
'    Set copyRange = FPwb.Sheets(1).Range(Cells(FPStartRow, 8).Address, Cells(FPEndRow, 11).Address)
'    copyRange.Copy Destination:=.Range(Cells(QAFPStartRow, 11).Address, Cells(QAFPEndRow, 14).Address)
'    Set copyRange = FPwb.Sheets(1).Range(Cells(FPStartRow, 14).Address, Cells(FPEndRow, 18).Address)
'    copyRange.Copy Destination:=.Range(Cells(QAFPStartRow, 17).Address, Cells(QAFPEndRow, 21).Address)

    'add time stamp and initial
    If .Cells(2, 2).Value < .Cells(3, 2).Value Then
        .range("B2:C2").Value = Split(Format(Now(), "mm/dd/yyyy") & "|" & QAiniTB.Value, "|")
    Else
        .range("B3:C3").Value = Split(Format(Now(), "mm/dd/yyyy") & "|" & QAiniTB.Value, "|")
    End If
End With

Call loginfo(siteName, "Flow Data and FP updated to " & RawEndDate)

LVFwb.Close savechanges:=False
Redwb.Close savechanges:=False
FPwb.Close savechanges:=False


'Set fso = Nothing
'Set copyRange = Nothing

'restart the screen flashing
Application.ScreenUpdating = True

End Sub

Private Sub QAsepFPbtn_Click()
'Hao Zhang @ 2015.2.4
'Seperates Field points into "historical" and "quarterly" in "Field Points Data" Tab

Call loadSite
'Dim srcWs As Worksheet, tgtWs As Worksheet
'Dim iRow As Integer, jRow As Integer, kRow As Integer
'
'For iSite = 0 To chosenSitesLB.ListCount - 1
'    If chosenSitesLB.Selected(iSite) = True Then
'        siteName = chosenSitesLB.List(iSite, 1)
'        Row = chosenSitesLB.List(iSite, 0)
'        Exit For
'    End If
'Next
'
'QtrYr = QAqtrCB.Value & "-" & Right(QAyrCB.Value, 2)
'Col = MyWs.Rows(1).Find(QtrYr).Column
'QAsheet = GetQApath(siteName, QtrYr)
'
'Set QAwb = Workbooks(fso.GetFileName(QAsheet))
Set srcws = QAwb.Worksheets("Site Info")
Set tgtws = QAwb.Worksheets("Field Points Data")

'clear target range contents first
With tgtws
    .range("A9:E196").ClearContents
    .range("G9:k46").ClearContents
End With

jRow = 9
kRow = 9
'set up the delimiter between "historical" and "new"
dtime = DateSerial(QAqtrYrCB.Column(1), QAqtrYrCB.Column(2), 1)
With srcws
    EOR = .Cells(.Rows.count, "C").End(xlUp).Row
    For iRow = 16 To EOR
        If .Cells(iRow, "B") < dtime Then
            tgtws.Cells(jRow, "A").Value = .Cells(iRow, "B").Value
            tgtws.Cells(jRow, "B").Value = .Cells(iRow, "H").Value
            tgtws.Cells(jRow, "C").Value = .Cells(iRow, "J").Value
            tgtws.Cells(jRow, "D").Value = .Cells(iRow, "K").Value
            tgtws.Cells(jRow, "E").Value = .Cells(iRow, "R").Value
            jRow = jRow + 1
        Else
            tgtws.Cells(kRow, "G").Value = .Cells(iRow, "B").Value
            tgtws.Cells(kRow, "H").Value = .Cells(iRow, "H").Value
            tgtws.Cells(kRow, "I").Value = .Cells(iRow, "J").Value
            tgtws.Cells(kRow, "J").Value = .Cells(iRow, "K").Value
            tgtws.Cells(kRow, "K").Value = .Cells(iRow, "R").Value
            kRow = kRow + 1
        End If
    Next
End With

Call loginfo(siteName, fso.GetFileName(QAsheet) & " FP seperated")


End Sub


Private Sub QAtrimTailBtn_Click()
'Hao Zhang @2015.2.6
'[trim tail]
'delete formula in empty cells so charts can be shown properly

Call loadSite

'For iSite = 0 To chosenSitesLB.ListCount - 1
'    If chosenSitesLB.Selected(iSite) = True Then
'        siteName = chosenSitesLB.List(iSite, 1)
'        Row = chosenSitesLB.List(iSite, 0)
'        Exit For
'    End If
'Next
'
'QtrYr = QAqtrCB.Value & "-" & Right(QAyrCB.Value, 2)
'Col = MyWs.Rows(1).Find(QtrYr).Column
'QAsheet = GetQApath(siteName, QtrYr)
'
'Set QAwb = Workbooks(fso.GetFileName(QAsheet))

With QAwb.Worksheets("Flow Data")
    'first and last column to be trimed
    Col_1 = .range("10:20").Find("Corrected Level").Column
    Col_4 = .range("10:20").Find("Corrected Temperature").Column
    'first and last row to be trimed
    Row_1 = .Cells(.Rows.count, 2).End(xlUp).Row + 1
    Row_4 = .Cells(.Rows.count, 1).End(xlUp).Row
    
    'abandoned because it's too slow
'    For Each Cell In .Range(.Cells(Row_1, Col_1).Address, .Cells(Row_4, Col_4).Address)
'        If Cell.HasFormula = True And Cell.Value = "" Then
'            Cell.ClearContents
'        End If
'    Next

    .range(.Cells(Row_1, Col_1).Address, .Cells(Row_4, Col_4)).ClearContents
End With

Call loginfo(siteName, fso.GetFileName(QAsheet) & " tail trimed")


End Sub



Private Sub QAappRainBtn_Click()
'Hao Zhang @ 2015.2.1
'pull rainfall data from PWD2010.mdb to Rainfall tab in QA sheet

Dim RG As Integer
Dim startTime As Date, endTime As Date
Dim wb As Workbook
Dim ws As Worksheet
Dim Cn As ADODB.Connection, rs As ADODB.Recordset
Dim MyConn, sSQL As String
Dim QAfilePath As String
Dim fd As Worksheet
Dim rainCol As Integer
'Dim vArr() As Variant
Dim BOR As Integer
Dim EOR As Long


'set RG automatically from QA sheet
''RG = RGCB.Value

startTime = StartTimeTB.Value
endTime = EndTimeTB.Value

Call loadSite

'For iSite = 0 To chosenSitesLB.ListCount - 1
'    If chosenSitesLB.Selected(iSite) = True Then
'        siteName = chosenSitesLB.List(iSite, 1)
'        Row = chosenSitesLB.List(iSite, 0)
'    End If
'Next
'
'QtrYr = QAqtrCB.Value & "-" & Right(QAyrCB.Value, 2)
'Col = MyWs.Rows(1).Find(QtrYr).Column
'QAfilePath = GetQApath(siteName, QtrYr)
'
''open the site QA sheet
'If fso.FileExists(QAfilePath) = True Then
'    Application.Workbooks.Open fileName:=QAfilePath
'Else
'    a = MsgBox("Sorry, can't find the QA file, please specify:", vbOKCancel, "Warning")
'    If a = vbOK Then
'        QAfilePath = GetFile("C:\")
'        Application.Workbooks.Open fileName:=QAfilePath
'    Else
'        Exit Sub
'    End If
'End If
'
'Set QAwb = Workbooks(fso.GetFileName(QAfilePath))
With QAwb
    Set fd = .Sheets("Flow Data")
    Set si = .Sheets("Site Info")
End With

'if RG# is declared explicitly, then use the input RG#
'otherwise, look in the QA sheet for RG#
If RGCB.Value = "" Then
    a = si.range("A1:C18").Find("RG #:").Offset(0, 1).Value
    If 1 <= a <= 35 Then
        RG = a
    Else
        RG = InputBox("Which Rain Gauge? (1-35)", "Select")
    End If
Else
     RG = RGCB.Value
End If
'set the worksheet if exist, otherwise, create one then set it as ws

'If QAwb.Sheets("Rainfall Data") Is Nothing Then
'    If QAwb.Sheets("Rainfall") Is Nothing Then

If TabExists("Rainfall Data", QAwb) = False Then
    If TabExists("Rainfall", QAwb) = False Then
        Set ws = QAwb.Sheets.Add(after:=Sheets(QAwb.Sheets.count))
        ws.Name = "Rainfall Data"
    Else
        Set ws = QAwb.Sheets("Rainfall")
        ws.Name = "Rainfall Data"
    End If
Else
    Set ws = QAwb.Sheets("Rainfall Data")
End If

 'Set source
MyConn = "C:\Rainfall\PWDRAIN2010\PWDRAIN2010.mdb"
 'Create query
sSQL = "SELECT Daytime, finalRG" & RG & " FROM [FinalAll(" & Year(startTime) & ")] WHERE (((Daytime) >= #" & startTime & "# And (Daytime) < #" & endTime & "#));"

 'Create RecordSet
Set Cn = New ADODB.Connection
With Cn
    .Provider = "Microsoft.ACE.OLEDB.12.0"  'ACE is a newer and better oleDB driver than JET
   '.Provider = "Microsoft.Jet.OLEDB.4.0"
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

'link cells in Flow Data and Rainfall Data
'bypass the 2min or 5min data since the time steps won't be consistent
If Not siteName Like "*_*min*" Then
    'find the column for rainfall in Flow Data
    With fd
        rainCol = .range("A12:AZ13").Find("Rain Fall Data").Column
        'convert the column number back to letter
        vArr = Split(.Cells(1, rainCol).Address(True, False), "$")
        'find the first row of data (BOR=begin of row)
        BOR = .range("A12:AZ13").Find("Rain Fall Data").Row + 2
        'find the last row of data (EOR=end of row)
        EOR = .Cells(.Rows.count, "A").End(xlUp).Row
        'link Flow data!rainfall data (column AC) to Rainfall!rainfall(column B)
        .range(vArr(0) & BOR, vArr(0) & EOR).Formula = "='Rainfall Data'!B2"
    End With
End If

're-initialize RG#
RGCB.Value = ""

Call loginfo(siteName, fso.GetFileName(QAsheet) & " Rainfall Imported")


End Sub
Private Sub loginfo(site, msg)
'Hao Zhang @ 2015.2.9
'log operations in the QA Logbook.xlsm
'identify the log sheet
Dim logFl As String
Dim logwb As Workbook

logFl = "C:\Users\hao.zhang\Desktop\QA logbook.xlsm"
'determine if logbook.xlsm is open
If IsWorkBookOpen(logFl) = True Then
    Set logwb = Workbooks(Dir(logFl))
Else
    Set logwb = Workbooks.Open(fileName:=logFl)
End If

'identify the log tab
Dim logtab As String
Dim logws As Worksheet

logtab = Format(Now(), "MMMYY") & "Log"
'determine if logtab exists already
If TabExists(logtab, logwb) = True Then
    Set logws = logwb.Worksheets(logtab)
Else
    Set logws = logwb.Sheets.Add(after:=Sheets(logwb.Sheets.count))
    With logws
        .Name = logtab
        .range("A1:D1").Value = Split("No|Date|Site|Action", "|")
        .Columns(1).Width = 5
        .Columns(2).Width = 20
        .Columns(3).Width = 10
        .Columns(4).Width = 40
    End With
End If

'log the action in the first available row in logbook
With logws
    logRow = .Cells(.Rows.count, 1).End(xlUp).Row + 1
    .Cells(logRow, 1).Value = logRow - 1
    .Cells(logRow, 2).Value = Now()
    .Cells(logRow, 3).Value = site
    .Cells(logRow, 4).Value = msg
End With

'adjust windows for visual check
With logwb.Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth * 0.5
    .Height = Application.UsableHeight * 0.2
    .Top = Application.UsableHeight * 0.8
    .Left = 0
    .ScrollColumn = 1
    If logRow > 10 Then
        .ScrollRow = logRow - 2
    End If
End With

logwb.Save

End Sub


Private Sub QAComboBtn_Click()
'Hao Zhang @2015.2.9
'A combination of QAQC elaborations
'create a QA sheet (if needed), import site info, import Flow data, Field points, rainfall, seperate FP points, trim tails.


Call QAappFlowBtn_Click
DoEvents
Call QAappRainBtn_Click
DoEvents
Call QAsepFPbtn_Click
DoEvents
Call QAtrimTailBtn_Click
DoEvents
Call MoveDoneSiteBtn_Click
DoEvents

'loginfo
With QAwb.Sheets("Flow Data")
    QAupdDate = Format(.Cells(.Rows.count, 2).End(xlUp).Offset(0, -1).Value, "YYYY/MM/DD")
End With

Call loginfo(siteName, "QA sheet updated to " & QAupdDate)

End Sub

Private Sub QAadjChtBtn_Click()
'Hao Zhang @ 2015.2.11
'adjust charts data range and time scale

'Dim wb As Workbook

'srcPath = QAtempTB.Text

'
''enter the time interval in minutes
'intvl = QAtempIntvlTB.Text
'
''use the fso method to get the dirctory
'flPath = fso.GetParentFolderName(srcPath) & "\"
'srcFile = fso.GetBaseName(srcPath)
'srcExt = "." & fso.GetExtensionName(srcPath)
'tgtName = "QAQC template (" & mQtr & "-" & Right(mYear, 2) & ")_" & intvl & "min"
''get an unique file name using function GetUniqueName()
'tgtPath = GetUniqueName(flPath & tgtName & srcExt)
'
''duplicate source file, rename the generated file.
'If fso.FileExists(srcPath) = True Then
'    'if sourcefile and target file are same, then not duplicate is needed
'    If srcPath <> tgtPath Then
'    FileCopy srcPath, tgtPath
'    End If
'Else
'    MsgBox "the source template doesn't exist."
'    Exit Sub
'End If
'
''Hao Zhang @ 2015.1.30
''opens the new template file
'Application.Workbooks.Open fileName:=tgtPath
'Set wb = Workbooks(fso.GetFileName(tgtPath))


'stop screen flashing
Application.ScreenUpdating = False

Call loadSite

mYear = QAqtrYrCB.Column(1)
startMonth = QAqtrYrCB.Column(3)


'set up the StartDate and endDate, endDate will be updated for each month in functions
month1 = DateSerial(mYear, startMonth, 1)
Month2 = DateAdd("m", 1, month1)
Month3 = DateAdd("m", 2, month1)
Month4 = DateAdd("m", 3, month1)
'identify the row number of the begining of each month (Row1,2,3) and the end of the last month(Row4)
'fixed the bug that Row numbers are not integers
Dim Row1 As Long
Dim Row2 As Long
Dim Row3 As Long
Dim Row4 As Long

Row1 = QAwb.Worksheets("Flow Data").range("A1:A40").Find("DateTime").Row + 2
Row2 = Row1 + ((Day(Month2 - 1)) * 60 / intvl * 24)
Row3 = Row2 + ((Day(Month3 - 1)) * 60 / intvl * 24)
Row4 = Row3 + ((Day(Month4 - 1)) * 60 / intvl * 24)

''update the month-year in table
'With wb.Sheets("Site Info")
'    .Range("E5").Value = Format(Month1, "mmmm yyyy")
'    .Range("E6").Value = Format(Month2, "mmmm yyyy")
'    .Range("E7").Value = Format(Month3, "mmmm yyyy")
'    .Range("E8").Value = mQtr & "-" & mYear
'End With
'
''correct an anormaly in tab name
'If TabExists("QAQC  Notes", wb) = True Then
'    Sheets("QAQC  Notes").Name = "QAQC Notes"
'End If
'
''update the month-year in table
'With wb.Sheets("QAQC Notes")
'    .Range("A7").Value = Format(Month1, "mmmm yyyy")
'    .Range("A8").Value = Format(Month2, "mmmm yyyy")
'    .Range("A9").Value = Format(Month3, "mmmm yyyy")
'End With
'
'With wb.Worksheets("Flow Data")
'    'clear contents of previous dates
'    .Range("A" & Row1, .Cells(.Rows.Count, "A").End(xlUp)).ClearContents
'    'fill range (fastest way)
'    'for comparision of other methods, refer to [populate_time_range] test()
'    .Range("A" & Row1).Value = Month1
'    .Range("A" & Row1 + 1, "A" & Row4).Formula = "=A" & Row1 & "+(" & intvl & "/60)/24"
'    .Range("A" & Row1, "A" & Row4).NumberFormat = "mm/dd/yyyy hh:mm:ss"
'
''update the month-year in the Percent Recovery in Flow Data
'    .Range("H5").Value = Format(startDate, "mmmm")
'    .Range("H6").Value = Format(DateAdd("m", 1, startDate), "mmmm")
'    .Range("H7").Value = Format(DateAdd("m", 2, startDate), "mmmm")
''change the Percent Recovery in Flow Data
'    .Range("I5").Value = "=(COUNT(U" & Row1 & ":U" & (Row2 - 1) & "))/(COUNT(A" & Row1 & ":A" & (Row2 - 1) & "))"
'    .Range("I6").Value = "=(COUNT(U" & Row2 & ":U" & (Row3 - 1) & "))/(COUNT(A" & Row2 & ":A" & (Row3 - 1) & "))"
'    .Range("I7").Value = "=(COUNT(U" & Row3 & ":U" & (Row4 - 1) & "))/(COUNT(A" & Row3 & ":A" & (Row4 - 1) & "))"
'    .Range("j5").Value = "=(COUNT(w" & Row1 & ":w" & (Row2 - 1) & "))/(COUNT(A" & Row1 & ":A" & (Row2 - 1) & "))"
'    .Range("j6").Value = "=(COUNT(w" & Row2 & ":w" & (Row3 - 1) & "))/(COUNT(A" & Row2 & ":A" & (Row3 - 1) & "))"
'    .Range("j7").Value = "=(COUNT(w" & Row3 & ":w" & (Row4 - 1) & "))/(COUNT(A" & Row3 & ":A" & (Row4 - 1) & "))"
'    .Range("k5").Value = "=(COUNT(v" & Row1 & ":w" & (Row2 - 1) & "))/(COUNT(A" & Row1 & ":A" & (Row2 - 1) & "))"
'    .Range("k6").Value = "=(COUNT(v" & Row2 & ":w" & (Row3 - 1) & "))/(COUNT(A" & Row2 & ":A" & (Row3 - 1) & "))"
'    .Range("k7").Value = "=(COUNT(v" & Row3 & ":w" & (Row4 - 1) & "))/(COUNT(A" & Row3 & ":A" & (Row4 - 1) & "))"
'End With

'Part II: Charts

''change Charts names
'With wb
'    For i = 1 To 18
'    .Charts(i).Name = "chart" & i
'    Next i
'    j = 1
'    For i = 0 To 2
'        .Charts(j).Name = MonthName(StartMonth + i, True) & " SP (Flow)"
'        .Charts(j + 1).Name = MonthName(StartMonth + i, True) & " SP CORR (Flow)"
'        .Charts(j + 2).Name = MonthName(StartMonth + i, True) & " SP (Vel)"
'        .Charts(j + 3).Name = MonthName(StartMonth + i, True) & " SP CORR (Vel)"
'        .Charts(j + 4).Name = MonthName(StartMonth + i, True) & " TS"
'        .Charts(j + 5).Name = MonthName(StartMonth + i, True) & " TS CORR"
'        j = j + 6
'    Next i
'End With

' change the monthly charts' data range and scale

Call QAQC_SP(month1, Month2, Row1, (Row2 - 1))
Call QAQC_TS(month1, Month2, Row1, (Row2 - 1), intvl)

Call QAQC_SP(Month2, Month3, Row2, (Row3 - 1))
Call QAQC_TS(Month2, Month3, Row2, (Row3 - 1), intvl)

Call QAQC_SP(Month3, Month4, Row3, (Row4 - 1))
Call QAQC_TS(Month3, Month4, Row3, (Row4 - 1), intvl)

' change the ALL charts data range and scale
Call QAQC_SP("ALL", Month4, Row1, (Row4 - 1))
Call QAQC_TS("ALL", Month4, Row1, (Row4 - 1), intvl)
'change the rest 4 charts' data range only, scales are not changed
Call QAQC_misc(Row1, (Row4 - 1))

'save the generated new template
QAwb.Save
''unload the wb object
'Set wb = Nothing
Application.ScreenUpdating = True

Call loginfo(siteName, "Charts data range and time scale adjusted")

End Sub


Private Sub IntactRptBtn_Click()
'Hao Zhang @ 2014.2.13
'create a report for site intactness
'After digging into FTP protocols, it is not feasible to compare with FTP server
'as a surrogate, this sub will only check the directory on PWD servers
'and records any missing raw folders for a given time range (+/- 3days)

'use the QA logbook.xlsm as the output
Dim flg As Integer
Dim rawDate As String

rawDate = Format(CDate(Right(RawDateTB, 10)), "YYYY-MM-DD")

For i = 2 To SiteCount + 1
    sitePath = MyWs.Cells(i, siteFldr_Col).Value
    siteName = MyWs.Cells(i, siteName_Col).Value
    'dealing with naming variations
    If fso.FolderExists(sitePath & "\Raw Data\") = True Then
        rawPath = sitePath & "\Raw Data\"
    Else
        rawPath = sitePath & "\RawData\"
    End If
    flg = 0
    'find the +/- 5 days raw data
    For j = -5 To 5
        rawFldr = "PE " & Format(DateAdd("d", j, rawDate), "YYYY-MM-DD")
        If fso.FolderExists(rawPath & rawFldr) = True Then
            flg = flg + 1
            Exit For
        End If
    Next j
    If flg = 0 Then
        Call loginfo(siteName, RawDateTB & " folder missing")
    End If
DoEvents
Next i

Call loginfo("all sites", "Raw Folder intactness check completed for " & RawDateTB)

End Sub

Private Sub QAadjYaxisBtn_Click()
'Hao Zhang @ 2015.2.14
'change the primary/secondary axis groups in TS charts based on max values

'''not tested yet

'get site info
Call loadSite

'get average value of level, velocity, and flow
With QAwb.Worksheets("Flow Data")
    Lvl_Col = .range("10:25").Find("Level 1", lookat:=xlWhole).Column
    Vel_Col = .range("10:25").Find("Vel 1", lookat:=xlWhole).Column
    Flw_Col = .range("10:25").Find("Flow 1").Column
    startRow = .range("10:25").Find("Flow 1").Row + 2
    endRow = .Cells(.Rows.count, Lvl_Col).End(xlUp).Row
    'bypass blank sheet condition
    If startRow < endRow Then
        'get median level, velocity and flow
        aveLvl = Application.WorksheetFunction.Median(.range(.Cells(startRow, Lvl_Col).Address & ":" & .Cells(endRow, Lvl_Col).Address))
        aveVel = Application.WorksheetFunction.Median(.range(.Cells(startRow, Vel_Col).Address & ":" & .Cells(endRow, Vel_Col).Address))
        aveFlw = Application.WorksheetFunction.Median(.range(.Cells(startRow, Flw_Col).Address & ":" & .Cells(endRow, Flw_Col).Address))
        'get the ratio of each pair, and values are always larger than 1
        L_V = IIf(aveLvl / aveVel > 1, aveLvl / aveVel, aveVel / aveLvl)
        V_F = IIf(aveVel / aveFlw > 1, aveVel / aveFlw, aveFlw / aveVel)
        L_F = IIf(aveLvl / aveFlw > 1, aveLvl / aveFlw, aveFlw / aveLvl)
    Else
        MsgBox "Cannot process blank sheets, please fill in data first and try again."
        Exit Sub
    End If
End With

'revised @ 2015.2.26
'find the closest pair of parameters
MinRatio = Application.WorksheetFunction.Min(L_V, V_F, L_F)

Select Case MinRatio
    Case Is = V_F
        'condition 1: Primary axis:Flow & Vel, Secondary axis: Level (default)
        flw = 1
        vel = 1
        lvl = 2
        y1 = "Flow (MGD)/Velocity (fps)"
        y2 = "Level (in)"
    Case Is = L_F
        'condition 2: Primary axis:Vel, Secondary axis: Flow & Level
        flw = 2
        vel = 1
        lvl = 2
        y1 = "Velocity (fps)"
        y2 = "Flow (MGD)/Level (in)"
    Case Is = L_V
        'condition 3: Primary axis:Flow, Secondary axis: Level & Vel
        flw = 1
        vel = 2
        lvl = 2
        y1 = "Flow (MGD)"
        y2 = "Level (in)/Velocity (fps)"
End Select

''enumerate all possible combinations, some condition could be combined
''but intentially retained to better demostrating the logic process
'If L_V > 10 Then
'    If V_F > 10 Then
'        If L_F > 10 Then
'            If V_F > L_F Then
'            'put LF together
'                cond = 2
'            Else
'            'put VF together
'                cond = 1
'            End If
'        Else
'            'put LF together
'            cond = 2
'        End If
'    Else
'        If L_F > 10 Then
'            'put VF together
'            cond = 1
'        Else
'            If V_F > L_F Then
'            'put LF together
'                cond = 2
'            Else
'            'put VF together
'                cond = 1
'            End If
'        End If
'    End If
'Else
'    If V_F > 10 Then
'        If L_F > 10 Then
'            'put LV together
'            cond = 3
'        Else
'            If L_V > L_F Then
'            'put LF together
'                cond = 2
'            Else
'            'put LV together
'                cond = 3
'            End If
'        End If
'    'all similar condition
'    Else
'        If L_F > 10 Then
'            If V_F > L_V Then
'            'put LV together
'                cond = 3
'            Else
'            'put VF together
'                cond = 1
'            End If
'        Else
'            If V_F > L_V Then
'                If V_F > L_F Then
'                'put LF together
'                    cond = 2
'                Else
'                'put LV together
'                    cond = 3
'                End If
'            Else
'                'put VF together
'                cond = 1
'            End If
'        End If
'    End If
'End If


For Each Cht In QAwb.Charts
    'change TS plots
    If Cht.Name Like "*TS" Then
        With Cht
            'change primary/seconary axis group
            .SeriesCollection("Flow 1").AxisGroup = flw
            .SeriesCollection("Flow 2").AxisGroup = flw
            .SeriesCollection("Field Flow").AxisGroup = flw
            .SeriesCollection("Vel 1").AxisGroup = vel
            .SeriesCollection("Vel 2").AxisGroup = vel
            .SeriesCollection("Field Velocity").AxisGroup = vel
            .SeriesCollection("Level 1").AxisGroup = lvl
            .SeriesCollection("Level 2").AxisGroup = lvl
            .SeriesCollection("Field Level").AxisGroup = lvl
            .SeriesCollection("Silt").AxisGroup = lvl
            'change primary/seconary axis titles
            .Axes(xlValue, xlPrimary).AxisTitle.Text = y1
            .Axes(xlValue, xlPrimary).MaximumScaleIsAuto = True
            .Axes(xlValue, xlSecondary).AxisTitle.Text = y2
            .Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True
        End With
    End If
    'change TS CORR plots, except ALL TS CORR
    If Cht.Name Like "*TS CORR" And Cht.Name <> "ALL TS CORR" Then
        With Cht
            'change primary/seconary axis group
            .SeriesCollection("Flow 1").AxisGroup = 1
            .SeriesCollection("Field Flow").AxisGroup = 1
            .SeriesCollection("Level 1").AxisGroup = 2
            .SeriesCollection("Field Level").AxisGroup = 2
            'change primary/seconary axis titles
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Flow (MGD)"
            .Axes(xlValue, xlPrimary).MaximumScaleIsAuto = True
            .Axes(xlValue, xlSecondary).AxisTitle.Text = "Level (in)"
            .Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True
        End With
     End If
     'change ALL TS CORR
     If Cht.Name = "ALL TS CORR" Then
        With Cht
        'change primary/seconary axis group
            .SeriesCollection("Corrected Flow (mgd)").AxisGroup = 1
            .SeriesCollection("Field Flow").AxisGroup = 1
            .SeriesCollection("Corrected Level (in)").AxisGroup = 2
            .SeriesCollection("Field Level").AxisGroup = 2
            'change primary/seconary axis titles
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Flow (MGD)"
            .Axes(xlValue, xlPrimary).MaximumScaleIsAuto = True
            .Axes(xlValue, xlSecondary).AxisTitle.Text = "Level (in)"
            .Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True
        End With
     End If
Next
    
End Sub
'


'Private Sub QArefDataBtn_Click()
''Stop the screen from flashing
'Application.ScreenUpdating = False
''Initial values
'flag_type = ""
'Dim ElementID As Long: ElementID = xlSeries
'
'With ActiveWorkbook.ActiveChart
'    .GetChartElement x, y, ElementID, ByVal Arg1, ByVal Arg2
''If ElementID = xlSeries Then
'    If Arg2 > 0 Then
'
'        x = WorksheetFunction.Index(.SeriesCollection(Arg1).XValues, Arg2)
'        Adj_Arg2 = Arg2 + 15 - 1 ' 1 for headers
'        iReply = MsgBox(Prompt:="x = " & x & "? Navigate to the data row?", _
'                            Buttons:=vbYesNo, Title:="Locate data row?")
'            If iReply = vbYes Then
'                flag_type = "flag_goto"
'            End If
'            If iReply = vbNo Then
'                flag_type = ""
'            End If
'    End If
''End If
'End With
''Goto Section
'
'If flag_type = "flag_goto" Then
'        'Activate the flow data worksheet, go to the row where the x:y data point occurents (Arg2 of plot)
'        Sheets("Flow Data").Activate
'        ActiveWindow.ScrollRow = Adj_Arg2 - 10
'        Cells(Adj_Arg2, 1).Select
'End If
'
'End Sub

Private Sub QAfillYlwShtBtn_Click()
'Hao Zhang @ 2015.3.23
'fill the %completed to the yellow sheet
'*********need test*******************
'ask the user to specify the yellow sheet
'use a global variable for the path of ylwsheet, so user only need to specify it once
'shared with BigPicturePPT
Dim iMon As Integer
Dim iYear As Integer
iMon = BigPpptMonCB.Text
iYear = BigPpptYrCB.Text

Dim ylwWb As Workbook
If ylwsheet = "" Then
    If MsgBox("Please select the yellow sheet you want to refer to:", vbOKCancel) = vbOK Then
        ylwPath = "M:\Data\Temporary Monitors\Flow Monitoring\Supplementary Documents\CSL Meeting Notes\" & iYear & "\" & MonthName(iMon + 1) & " " & iYear & "\"
        ylwsheet = GetFile(ylwPath)
    Else
        'cancel the operation
        Exit Sub
    End If
End If

If ylwsheet <> "" Then
    If IsWorkBookOpen(ylwsheet) = False Then
        Set ylwWb = Workbooks.Open(fileName:=ylwsheet)
    Else
        Set ylwWb = Workbooks(Dir(ylwsheet))
    End If
Else
    MsgBox "No file is selected."
    Exit Sub
End If

Dim QAref As String

'get table defination of the yellow sheet
With ylwWb.Sheets(1)
    Site_Col = .range("1:10").Find("OOW RDII Monitoring Sites:").Column
    Recov_Col = .range("1:5").Find("Percent Recovery").Column
    RDII_Row = .Columns(Site_Col).Find("OOW RDII Monitoring Sites:").Row
    DCIA_Row = .Columns(Site_Col).Find("OOW DCIA Monitoring Sites:").Row
    SWM_Row = .Columns(Site_Col).Find("Stormwater Monitoring Sites:").Row
    EOR = .Columns(Site_Col).Find("Overall Percent Recovery").Row
    For N = RDII_Row To EOR
        If (N > RDII_Row And N < DCIA_Row - 2) Or (N > DCIA_Row And N < SWM_Row - 2) Or (N > SWM_Row And N < EOR - 1) Then
            siteName = .Cells(N, Site_Col).Value
            QtrYr = "Q" & DatePart("q", DateAdd("m", -1, Now())) & "-" & Right(Year(DateAdd("m", -1, Now())), 2)
            QAsheet = MyWs.Cells(MyWs.Columns(siteName_Col).Find(siteName).Row, MyWs.Rows(1).Find(QtrYr).Column).Value
            QApath = fso.GetParentFolderName(QAsheet)
            QAfile = fso.GetFileName(QAsheet)
            QAsheet = "Flow Data"
            For p = 0 To 2
                QArng = .Cells(5 + p, iMon Mod 3 + 8).Address
                QAref = "='" & QApath & "\[" & QAfile & "]" & QAsheet & "'!" & QArng
                .Cells(N, Recov_Col + p).Formula = QAref
                .Cells(N, Recov_Col + p).Value = .Cells(N, Recov_Col + p).Value
            Next
        End If
    Next
End With

End Sub




'*************************************************************************
'***********************Big Picture***************************************
'*************************************************************************
Private Sub BigPBurstCkB_Click()
'Hao Zhang @ 2015.1.29
'add a caution message for using burst mode in QA sheet importing
If BigPBurstCkB.Value = True Then
    a = MsgBox("The burst mode is not recommended, do you still want to enable it?" & Chr(13) & "You have to use 'Ctrl + Pause' to terminate the process", vbYesNo, "Really??")
    If a = vbYes Then
        BigPBurstCkB = True
        For isite = 0 To chosenSitesLB.ListCount - 1
            chosenSitesLB.Selected(isite) = True
        Next
    Else
        BigPBurstCkB = False
    End If
End If
End Sub

Private Sub BigPbrowse_Click()
'Hao Zhang @ 2015.2.15
'allow user selecting a folder to save the Big Picture Input files

'use this variable to save user's selection
Dim tempPath As String
tempPath = GetFldr(sPath)
'if no path was specified, then retain the original value
If tempPath <> "" Then
    BigPpathTB.Value = tempPath
End If

End Sub

Private Sub BigPfldrCreateBtn_Click()
'Hao Zhang @ 2015.2.15
'allow user creating a sub-folder to save the Big Picture Input files
If fso.FolderExists(BigPpathTB.Value & "\BigPicture") = False Then
    BigPpathTB.Value = fso.CreateFolder(BigPpathTB.Value & "\BigPicture").path
Else
    MsgBox "The specified folder already exists."
End If

End Sub

Private Sub BigPinputGenBtn_Click()
'Hao Zhang @ 2015.2.15
'create Big Picture Input files in specified folder
'get the destination folder that saves all Big Picture input files in separate subfolders
Dim isite As Integer

'check if chosenSiteLB is not empty and at least one site was selected
If chosenSitesLB.ListCount > 0 And chosenSitesLB.ListIndex <> -1 Then
    For isite = 0 To chosenSitesLB.ListCount - 1
        If chosenSitesLB.Selected(isite) = True Then
            siteRow = chosenSitesLB.List(isite, 0)
            siteName = MyWs.Cells(siteRow, siteName_Col).Value
            If BigPappOB = True Then
                Call BigPappend(siteName)
            Else
                Call BigPcomplete(siteName)
            End If
            Call BigPsaveBtn_Click
         End If
         
         If chosenSitesLB.ListIndex = 0 Then
            'single mode
            Exit Sub
         End If
    DoEvents
    Next
Else
    MsgBox "please add sites to 'Chosen Sites' first", vbCritical, "Error"
    Exit Sub
End If

End Sub
Sub BigPappend(siteName As String)
'Hao Zhang originated @ 2014.11.21
'Hao Zhang revised @ 2015.2.15
'prepare datasheets for bigPicture analysis following steps below:
'1. copy most recent 'Combined_field_points.csv', 'CombinedQAQC.csv' to local drive in new names incl. date and initial
'2. copy 'input.csv' and 'Run_BigPicture.bat' to local drive
'3. loop through QA sheets from the appending point, append Field Point and Flow data to the CSVs
'4. update input file names and paths in input.csv
'keyword: create folder, copy & paste file, partial match file name
'(dir), move file, For-loop, do-while loop


'PART 0 : Initialization
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Goto ActiveCell, True

'get the destination folder to save generated files
Dim localpath As String
localpath = BigPpathTB.Value
'get the row number in CurSitesLB
siteRow = MyWs.Columns(siteName_Col).Find(siteName).Row
Dim ini As String
If IsNull(BigPiniTB.Text) = False Then
    ini = BigPiniTB.Text
Else
    ini = InputBox("Please enter your initial")
End If

'get source file folder, dealing with naming variations
Dim srcfldr As String
If fso.FolderExists(MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\BigPicture\") = True Then
    srcfldr = MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\BigPicture\"
ElseIf fso.FolderExists(MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\Big Picture\") = True Then
    srcfldr = MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\Big Picture\"
Else
    MsgBox "please specify the Big Picture folder path on server:"
    srcfldr = GetFldr(MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\")
End If

'find the first file that partially matches the source file name
'only the filename with extension is assigned to sFound1 and sFound2
' * is used as a wildcard
'set the full path of source files
Dim sFound1 As String
Dim source1 As String
sFound1 = Dir(srcfldr & siteName & "_Combined_field_points*.csv")
If sFound1 <> "" Then
    source1 = srcfldr & sFound1
Else
    MsgBox "Please specify Combined field points file:"
    source1 = GetFile(srcfldr)
End If

Dim sFound2 As String
Dim source2 As String
sFound2 = Dir(srcfldr & siteName & "_CombinedQAQC*.csv")
If sFound2 <> "" Then
    source2 = srcfldr & sFound2
Else
    MsgBox "Please specify CombinedQAQC file:"
    source2 = GetFile(srcfldr)
End If

Dim source3 As String
If fso.FileExists(srcfldr & "input.csv") = True Then
    source3 = srcfldr & "input.csv"
Else
    MsgBox "Please specify input.csv file:"
    source3 = GetFile(srcfldr)
End If

'used the updated bat file from local drive instead (w/ updated R version & server name)
Dim source4 As String
Dim tempPath As String
source4 = localpath & "\" & "Run_BigPicture.bat"
    'If fso.FileExists(source4) = False Then
    '    MsgBox "Please specify the Run_BigPicture.bat file:"
    '    tempPath = GetFile("C:\")
    '    If tempPath <> "" Then
    '        fso.CopyFile tempPath, source4
    '    Else
    '        MsgBox "can't find the Run_BigPicture.bat file."
    '        Exit Sub
    '    End If
    'End If

If fso.FileExists(source4) = False Then
    'create the bat file directly
    ifileNum = FreeFile
    Open source4 For Output As #ifileNum
        Print #ifileNum, "@echo Disregard CMD.exe UNC path error. Handling location manually."
        Print #ifileNum, "@echo off"
        Print #ifileNum, "set calldir=%~dp0"
        Print #ifileNum, "Set scriptpath=" & Chr(34) & "\\pwdhqr\oows\Modeling\Data\Temporary Monitors\Flow Monitoring\QAQC Procedures_CDM\BigPicture\BigPicture_v1.5.R" & Chr(34)
        Print #ifileNum, "set temppath=%TEMP%\Big_Picture.R"
        Print #ifileNum, "Set rpath=" & Chr(34) & "C:\Program Files\R\R-3.0.2\bin\Rscript.exe" & Chr(34)
        Print #ifileNum, "copy %scriptpath% %temppath%"
        Print #ifileNum, "pushd %calldir%"
        Print #ifileNum, "%rpath% %temppath%"
        Print #ifileNum, "pause"
    Close #ifileNum
End If

'set the full path of target files
Dim target1 As String
Dim target2 As String
Dim target3 As String
Dim target4 As String
target1 = localpath & "\" & siteName & "\" & siteName & "_Combined_field_points_" & Format(Now(), "yymmdd") & "_" & ini & ".csv"
target2 = localpath & "\" & siteName & "\" & siteName & "_CombinedQAQC_" & Format(Now(), "yymmdd") & "_" & ini & ".csv"
target3 = localpath & "\" & siteName & "\" & "input.csv"
target4 = localpath & "\" & siteName & "\" & "Run_BigPicture.bat"


'PART I: File Handling
'if site subfolder is not exist, create the folder first
If fso.FolderExists(localpath & "\" & siteName & "\") = False Then
    fso.CreateFolder localpath & "\" & siteName & "\"
End If
'copy/overwrite Combined_FP from server to local drive
fso.CopyFile source1, target1
'copy/overwrite Combined_FP from server to local drive
fso.CopyFile source2, target2
'copy/overwrite input from server to local hard drive
fso.CopyFile source3, target3
'copy/overwrite Run_bigPicture from BigPicutre folder to local hard drive
fso.CopyFile source4, target4


'PART II : add data to FP file (si->FPwb): get latest QA sheet, copy the difference

'open FP file on local drive
Dim FPwb As Workbook
If IsWorkBookOpen(target1) = False Then
    Set FPwb = Workbooks.Open(fileName:=target1)
Else
    Set FPwb = Workbooks(fso.GetFileName(QAsheet))
End If

'adjust windows for visual check
With FPwb.Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth * 0.5
    .Height = Application.UsableHeight * 0.6
    .Top = 0
    .Left = 0
    .ScrollRow = 1
    .ScrollColumn = 1
End With

Dim FPws As Worksheet
Set FPws = FPwb.Sheets(1)
FPws.Columns(1).ColumnWidth = 18
With FPws.range("1:1")
    FP_dTime_Col = .Find("dtime").Column
    FP_level_Col = .Find("level").Column
    FP_flow_Col = .Find("flow").Column
    FP_velocity_Col = .Find("velocity").Column
    FP_start_Row = .Find("dtime").Row + 1
    FP_end_Row = FPws.Cells(FPws.Rows.count, FP_dTime_Col).End(xlUp).Row
    FP_end_Date = FPws.Cells(FP_end_Row, FP_dTime_Col).Value
End With

'get the lastest QA sheet
QAsheet = MyWs.Cells(siteRow, endQA_Col).End(xlToLeft).Value

'open QA_sheet
If IsWorkBookOpen(QAsheet) = False Then
    Set QAwb = Workbooks.Open(fileName:=QAsheet, UpdateLinks:=False, ReadOnly:=True)
Else
    Set QAwb = Workbooks(fso.GetFileName(QAsheet))
End If

'adjust windows for visual check
'With QAwb.Windows(1)
'    .WindowState = xlNormal
'    .Width = Application.UsableWidth
'    .Height = Application.UsableHeight * 0.2
'    .Top = Application.UsableHeight * 0.8
'    .Left = 0
'    .ScrollColumn = 1
'    .ScrollRow = 1
'End With

Dim si As Worksheet
Set si = QAwb.Sheets("site info")
'find the columns number of dTime, field level, field flow, and field velocity
With si.range("10:25")
    si_dTime_Col = .Find("Date Time").Column
    si_FLvl_Col = .Find(" Field Level (inches)").Column
    si_FFlw_Col = .Find("Field Flow (mgd)").Column
    si_FVel_Col = .Find(" Field Velocity (fps)").Column
    si_Ftime_Col = .Find("Field Time").Column
    'si_start_Row = .Find("Date Time").Row + 1
    si_append_Row = si.Columns(si_dTime_Col).Find(Format(FP_end_Date, "m/d/yyyy h:mm"), LookIn:=xlValues, lookat:=xlWhole).Row + 1
    si_end_Row = si.Cells(si.Rows.count, si_Ftime_Col).End(xlUp).Row
End With

'copy new date
If si_end_Row > si_append_Row Then
    With FPws
        .range(.Cells(FP_end_Row + 1, FP_dTime_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_append_Row, FP_dTime_Col).Address).Value _
         = si.range(.Cells(si_append_Row, si_dTime_Col).Address, .Cells(si_end_Row, si_dTime_Col).Address).Value
        .range(.Cells(FP_end_Row + 1, FP_level_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_append_Row, FP_level_Col).Address).Value _
         = si.range(.Cells(si_append_Row, si_FLvl_Col).Address, .Cells(si_end_Row, si_FLvl_Col).Address).Value
        .range(.Cells(FP_end_Row + 1, FP_flow_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_append_Row, FP_flow_Col).Address).Value _
         = si.range(.Cells(si_append_Row, si_FFlw_Col).Address, .Cells(si_end_Row, si_FFlw_Col).Address).Value
        .range(.Cells(FP_end_Row + 1, FP_velocity_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_append_Row, FP_velocity_Col).Address).Value _
         = si.range(.Cells(si_append_Row, si_FVel_Col).Address, .Cells(si_end_Row, si_FVel_Col).Address).Value
        'force formating the dtime column
        .range(.Cells(FP_end_Row + 1, FP_dTime_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_append_Row, FP_dTime_Col).Address).NumberFormat = "m/d/yyyy h:mm"
    End With
End If

QAwb.Close savechanges:=False

'Part III: combined_QAQC (fd->FDws): find the appending point, then add QAQC data recursively
Dim FDwb As Workbook
If IsWorkBookOpen(target2) = False Then
    Set FDwb = Workbooks.Open(fileName:=target2)
Else
    Set FDwb = Workbooks(fso.GetFileName(target2))
End If

'adjust windows for visual check
With FDwb.Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth * 0.5
    .Height = Application.UsableHeight * 0.6
    .Top = 0
    .Left = Application.UsableWidth * 0.5
    .ScrollColumn = 1
    .ScrollRow = 1
    ''scroll to the end
    '.ScrollRow = FDwb.Sheets(1).Cells(1, 1).End(xlDown).Row - 10
End With

Dim FDws As Worksheet
Set FDws = FDwb.Sheets(1)
FDws.Columns(1).ColumnWidth = 18
'get table defination
With FDws.Rows(1)
    tgt_dTime_Col = .Find("dtime").Column
    tgt_level_Col = .Find("level").Column
    tgt_flow_Col = .Find("flow").Column
    tgt_velocity_Col = .Find("velocity").Column
    tgt_corr_flow_Col = .Find("corrected.flow").Column
    tgt_corr_level_Col = .Find("corrected.level").Column
    tgt_start_Row = .Find("dtime").Row + 1
End With

'get the appending point, determine the first QA sheet to be used
Dim tgt_end_Row As Long
Dim tgt_append_Date As Date

'so far Big Pictures are only based on 15 min data, but it can be refined
intvl = 15
tgt_end_Row = FDws.Cells(FDws.Rows.count, tgt_dTime_Col).End(xlUp).Row
tgt_end_Date = FDws.Cells(tgt_end_Row, tgt_dTime_Col).Value
'add a time step for appending point
'DateAdd(): "m"=month, "n"=minute,"yyyy"=year, "y"=day of the year, "w"=weekday, "ww"=week, "q"=quarter
tgt_append_Date = DateAdd("n", intvl, tgt_end_Date)
QtrYr = "Q" & DatePart("q", tgt_append_Date) & "-" & Right(DatePart("yyyy", tgt_append_Date), 2)
QtrYr_Col = MyWs.Rows(1).Find(QtrYr).Column

'get the lastest QA sheet
QAsheet = MyWs.Cells(siteRow, QtrYr_Col).Value

'loop through each QA sheet, appending data to current QAQC file
Do While QAsheet <> ""
    'open QA_sheet
    If IsWorkBookOpen(QAsheet) = False Then
        Set QAwb = Workbooks.Open(fileName:=QAsheet, UpdateLinks:=False, ReadOnly:=True)
    Else
        Set QAwb = Workbooks(fso.GetFileName(QAsheet))
    End If
    Set fd = QAwb.Sheets("Flow data")
    
    With QAwb.Windows(1)
        .WindowState = xlNormal
        .Width = Application.UsableWidth
        .Height = Application.UsableHeight * 0.2
        .Top = Application.UsableHeight * 0.8
        .Left = 0
        .ScrollColumn = 1
        .ScrollRow = 1
    End With
    'get table defination
    With fd.range("10:20")
        src_dtime_Col = .Find("DateTime").Column
        src_Lvl_Col = .Find("Level 1").Column
        src_Flw_Col = .Find("Flow 1").Column
        src_Vel_Col = .Find("Vel 1").Column
        src_Corr_Flw_Col = .Find("Corrected Flow").Column
        src_Corr_Lvl_Col = .Find("Corrected Level").Column
        'src_start_Row = .Find("Level 1").Row + 2
        src_append_Row = fd.Columns(src_dtime_Col).Find(Format(tgt_append_Date, "mm/dd/yyyy hh:mm:ss"), LookIn:=xlValues).Row
        src_end_Row = fd.Cells(fd.Rows.count, src_Lvl_Col).End(xlUp).Row
    End With
    
    'copy new date
    If src_end_Row > src_append_Row Then
        With FDws
            'for dates, direct passing formula-calculated value to target cell will cause rounding issue that every midnight is one day backwards
            'to bypass this issue, the pastespecial method is used and 'xlPasteValuesAndNumberFormats' argument is selected
            fd.range(.Cells(src_append_Row, src_dtime_Col).Address, .Cells(src_end_Row, src_dtime_Col).Address).Copy
            .range(.Cells(tgt_end_Row + 1, tgt_dTime_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, tgt_dTime_Col).Address).PasteSpecial xlPasteValuesAndNumberFormats
            '.Range(.Cells(tgt_end_Row + 1, tgt_dTime_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, tgt_dTime_Col).Address).Value _
             = fd.Range(.Cells(src_append_Row, src_dTime_Col).Address, .Cells(src_end_Row, src_dTime_Col).Address).Value
            .range(.Cells(tgt_end_Row + 1, tgt_level_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, tgt_level_Col).Address).Value _
             = fd.range(.Cells(src_append_Row, src_Lvl_Col).Address, .Cells(src_end_Row, src_Lvl_Col).Address).Value
            .range(.Cells(tgt_end_Row + 1, tgt_velocity_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, tgt_velocity_Col).Address).Value _
             = fd.range(.Cells(src_append_Row, src_Vel_Col).Address, .Cells(src_end_Row, src_Vel_Col).Address).Value
            .range(.Cells(tgt_end_Row + 1, tgt_flow_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, tgt_flow_Col).Address).Value _
             = fd.range(.Cells(src_append_Row, src_Flw_Col).Address, .Cells(src_end_Row, src_Flw_Col).Address).Value
            .range(.Cells(tgt_end_Row + 1, tgt_corr_flow_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, tgt_corr_flow_Col).Address).Value _
             = fd.range(.Cells(src_append_Row, src_Corr_Flw_Col).Address, .Cells(src_end_Row, src_Corr_Flw_Col).Address).Value
            .range(.Cells(tgt_end_Row + 1, tgt_corr_level_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, tgt_corr_level_Col).Address).Value _
             = fd.range(.Cells(src_append_Row, src_Corr_Lvl_Col).Address, .Cells(src_end_Row, src_Corr_Lvl_Col).Address).Value
            'force formating the dtime column
            .range(.Cells(tgt_end_Row + 1, tgt_dTime_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, tgt_dTime_Col).Address).NumberFormat = "m/d/yyyy h:mm"
        End With
    Else
        'break condition: when no more data could be added, exit the procedure
        Exit Do
    End If
    
    'FDwb.Save
        
    'get the new appending point, determine the next QA sheet to be used
    tgt_end_Row = FDws.Cells(FDws.Rows.count, tgt_dTime_Col).End(xlUp).Row
    tgt_end_Date = FDws.Cells(tgt_end_Row, tgt_dTime_Col).Value
    'add a time step for appending point
    tgt_append_Date = DateAdd("n", intvl, tgt_end_Date)
    QtrYr2 = "Q" & DatePart("q", tgt_append_Date) & "-" & Right(DatePart("yyyy", tgt_append_Date), 2)
    'if the new appending point is at a new Quarter sheet, then close the current one and loop to the new sheet
    'otherwise, exit the loop
    If QtrYr2 <> QtrYr Then
        QAwb.Close savechanges:=False
        QtrYr_Col = MyWs.Rows(1).Find(QtrYr2).Column
        'get the QA sheet
        QAsheet = MyWs.Cells(siteRow, QtrYr_Col).Value
    Else
        Exit Do
    End If
DoEvents
Loop

'update input.csv (step 5)
Workbooks.Open fileName:=target3
Cells(3, 2).Value = Dir(target1)
Cells(2, 2).Value = Dir(target2)
Cells(6, 2).Value = ""
Cells(7, 2).Value = ""
'adjust windows for visual check
With Workbooks("input.csv").Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth
    .Height = Application.UsableHeight * 0.2
    .Top = Application.UsableHeight * 0.6
    .Left = 0
    .ScrollRow = 1
    .ScrollColumn = 1
End With
'Workbooks("input.csv").Sheets(1).Columns(1).ColumnWidth = 25

'log operations

Call loginfo(siteName, "Big Picture input files generated (append)")

Application.DisplayAlerts = False
Application.ScreenUpdating = True

End Sub

Private Sub BigPcomplete(siteName As String)
'Hao Zhang @ 2015.2.17
'generate big picture input files
'create a site subfolder
'create 3 csv files, set file name and table defination
'get Run_BigPicture.bat, copy to destination folder
'fill FP file using the latest QA sheet
'fill QAQC file by looping through all available QA sheet
'fill input.csv with FP and QAQC name
'log action

'PART 0 : Initialization
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Goto ActiveCell, True

'get the destination folder to save generated files
Dim localpath As String
localpath = BigPpathTB.Value
siteRow = MyWs.Columns(siteName_Col).Find(siteName).Row

''get the row number in CurSitesLB
'Dim Row As Integer
'For isite = 0 To chosenSitesLB.ListCount - 1
'    If chosenSitesLB.Selected(isite) = True Then
'        Row = chosenSitesLB.List(isite, 0)
'        siteName = MyWs.Cells(Row, siteName_Col).Value
'        Exit For
'    End If
'Next

'set the full path of target files
Dim target1 As String
Dim target2 As String
Dim target3 As String
Dim target4 As String
tgtfldr = localpath & "\" & siteName & "\"
target1 = tgtfldr & siteName & "_Combined_field_points_" & Format(Now(), "yymmdd") & "_" & ini & ".csv"
target2 = tgtfldr & siteName & "_CombinedQAQC_" & Format(Now(), "yymmdd") & "_" & ini & ".csv"
target3 = tgtfldr & "input.csv"
target4 = tgtfldr & "Run_BigPicture.bat"


'PART I: File Handling

'cannot delete and create folders within the same sub
'regenerate site subfolder no matter if it exist or not
If fso.FolderExists(tgtfldr) = False Then
    fso.CreateFolder tgtfldr
Else
    'clear all contents in the tgtfldr
    fso.DeleteFile tgtfldr & "*"
End If

'set sheet numbers in new workbook to 1 (global variable)
Application.SheetsInNewWorkbook = 1
Set FPwb = Workbooks.Add
    With FPwb
        'add heading
        .Sheets(1).range("A1:D1") = Split("dtime|level|flow|velocity", "|")
        .Sheets(1).range("A:A").NumberFormat = "mm/dd/yyyy hh:mm"
        .Sheets(1).range("B:D").NumberFormat = "general"
        .SaveAs fileName:=target1, FileFormat:=xlCSV, CreateBackup:=False
    End With
    'adjust windows for visual check
    With FPwb.Windows(1)
        .WindowState = xlNormal
        .Width = Application.UsableWidth * 0.5
        .Height = Application.UsableHeight * 0.6
        .Top = 0
        .Left = 0
        .ScrollRow = 1
        .ScrollColumn = 1
    End With
    
    Dim FPws As Worksheet
    Set FPws = FPwb.Sheets(1)
    FPws.Columns(1).ColumnWidth = 18
    With FPws.range("1:1")
        FP_dTime_Col = .Find("dtime").Column
        FP_level_Col = .Find("level").Column
        FP_flow_Col = .Find("flow").Column
        FP_velocity_Col = .Find("velocity").Column
        FP_start_Row = .Find("dtime").Row + 1
        FP_end_Row = FPws.Cells(FPws.Rows.count, FP_dTime_Col).End(xlUp).Row
        FP_end_Date = FPws.Cells(FP_end_Row, FP_dTime_Col).Value
    End With

Set FDwb = Workbooks.Add
    With FDwb
        'add heading
        .Sheets(1).range("A1:F1") = Split("dtime|level|velocity|flow|corrected.flow|corrected.level", "|")
        .Sheets(1).range("A:A").NumberFormat = "mm/dd/yyyy hh:mm"
        .Sheets(1).range("B:E").NumberFormat = "general"
        .SaveAs fileName:=target2, FileFormat:=xlCSV, CreateBackup:=False
    End With
    'adjust windows for visual check
    With FDwb.Windows(1)
        .WindowState = xlNormal
        .Width = Application.UsableWidth * 0.5
        .Height = Application.UsableHeight * 0.6
        .Top = 0
        .Left = Application.UsableWidth * 0.5
        .ScrollColumn = 1
        .ScrollRow = 1
    End With
    Dim FDws As Worksheet
    Set FDws = FDwb.Sheets(1)
    FDws.Columns(1).ColumnWidth = 18
    'get table defination
    With FDws.Rows(1)
        tgt_dTime_Col = .Find("dtime").Column
        tgt_level_Col = .Find("level").Column
        tgt_velocity_Col = .Find("velocity").Column
        tgt_flow_Col = .Find("flow").Column
        tgt_corr_flow_Col = .Find("corrected.flow").Column
        tgt_corr_level_Col = .Find("corrected.level").Column
        tgt_start_Row = .Find("dtime").Row + 1
    End With

Set Inpwb = Workbooks.Add
    With Inpwb
        'add heading
        .Sheets(1).range("A1:A13") = Application.Transpose(Split("Parameter|timeseries.file|fieldpoints.file|timestep|filter.hours|y1max|y2max|xmin_scatter|xmax_scatter|ymin_velocity_scatter|ymax_velocity_scatter|ymin_flow_scatter|ymax_flow_scatter", "|"))
        .Sheets(1).range("B1") = "value"
        .Sheets(1).Cells(4, 2).Value = 15
        .Sheets(1).Cells(5, 2).Value = 4
        'may need to specify other value, if so, need to copy from previous file
        .SaveAs fileName:=target3, FileFormat:=xlCSV, CreateBackup:=False
    End With
    'adjust windows for visual check
    With Inpwb.Windows(1)
        .WindowState = xlNormal
        .Width = Application.UsableWidth
        .Height = Application.UsableHeight * 0.2
        .Top = Application.UsableHeight * 0.6
        .Left = 0
        .ScrollRow = 1
        .ScrollColumn = 1
    End With
    'Inpwb.Sheets(1).Columns(1).ColumnWidth = 25

'Dim tempPath As String
'If fso.FileExists(source4) = False Then
'    MsgBox "Please specify the Run_BigPicture.bat file:"
'    tempPath = GetFile("C:\")
'    If tempPath <> "" Then
'        fso.CopyFile tempPath, source4
'    Else
'        MsgBox "can't find the Run_BigPicture.bat file."
'        Exit Sub
'    End If
'End If
'
''copy/overwrite Run_bigPicture from BigPicutre folder to local hard drive
'fso.CopyFile source4, target4
'create the bat file
ifileNum = FreeFile
Open target4 For Output As #ifileNum
    Print #ifileNum, "@echo Disregard CMD.exe UNC path error. Handling location manually."
    Print #ifileNum, "@echo off"
    Print #ifileNum, "set calldir=%~dp0"
    Print #ifileNum, "Set scriptpath=" & Chr(34) & "\\pwdhqr\oows\Modeling\Data\Temporary Monitors\Flow Monitoring\QAQC Procedures_CDM\BigPicture\BigPicture_v1.5.R" & Chr(34)
    Print #ifileNum, "set temppath=%TEMP%\Big_Picture.R"
    Print #ifileNum, "Set rpath=" & Chr(34) & "C:\Program Files\R\R-3.0.2\bin\Rscript.exe" & Chr(34)
    Print #ifileNum, "copy %scriptpath% %temppath%"
    Print #ifileNum, "pushd %calldir%"
    Print #ifileNum, "%rpath% %temppath%"
    Print #ifileNum, "pause"
Close #ifileNum
    
    
'PART II : add data to FP file (si->FPwb): get latest QA sheet, copy the difference

'get the lastest QA sheet
For iCol = endQA_Col To startQA_Col Step -1
    If MyWs.Cells(siteRow, iCol).Value <> "" Then
        QAsheet = MyWs.Cells(siteRow, iCol).Value
        Exit For
    End If
Next

'open QA_sheet
If IsWorkBookOpen(QAsheet) = False Then
    Set QAwb = Workbooks.Open(fileName:=QAsheet, UpdateLinks:=False, ReadOnly:=True)
Else
    Set QAwb = Workbooks(fso.GetFileName(QAsheet))
End If

'adjust windows for visual check
'With QAwb.Windows(1)
'    .WindowState = xlNormal
'    .Width = Application.UsableWidth * 0.5
'    .Height = Application.UsableHeight
'    .Top = 0
'    .Left = Application.UsableWidth * 0.5
'    .ScrollColumn = 1
'    .ScrollRow = 1
'End With

Dim si As Worksheet
Set si = QAwb.Sheets("site info")
'find the columns number of dTime, field level, field flow, and field velocity
With si.range("10:25")
    si_dTime_Col = .Find("Date Time").Column
    si_FLvl_Col = .Find(" Field Level (inches)").Column
    si_FFlw_Col = .Find("Field Flow (mgd)").Column
    si_FVel_Col = .Find(" Field Velocity (fps)").Column
    si_Ftime_Col = .Find("Field Time").Column
    si_start_Row = .Find("Date Time").Row + 1
    'si_append_Row = si.Columns(si_dTime_Col).Find(Format(FP_end_Date, "m/d/yy h:mm"), LookIn:=xlValues, lookat:=xlWhole).Row + 1
    si_end_Row = si.Cells(si.Rows.count, si_Ftime_Col).End(xlUp).Row
End With

'copy new date
If si_end_Row > si_append_Row Then
    With FPws
        .range(.Cells(FP_end_Row + 1, FP_dTime_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_start_Row, FP_dTime_Col).Address).Value _
         = si.range(.Cells(si_start_Row, si_dTime_Col).Address, .Cells(si_end_Row, si_dTime_Col).Address).Value
        .range(.Cells(FP_end_Row + 1, FP_level_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_start_Row, FP_level_Col).Address).Value _
         = si.range(.Cells(si_start_Row, si_FLvl_Col).Address, .Cells(si_end_Row, si_FLvl_Col).Address).Value
        .range(.Cells(FP_end_Row + 1, FP_flow_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_start_Row, FP_flow_Col).Address).Value _
         = si.range(.Cells(si_start_Row, si_FFlw_Col).Address, .Cells(si_end_Row, si_FFlw_Col).Address).Value
        .range(.Cells(FP_end_Row + 1, FP_velocity_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_start_Row, FP_velocity_Col).Address).Value _
         = si.range(.Cells(si_start_Row, si_FVel_Col).Address, .Cells(si_end_Row, si_FVel_Col).Address).Value
        'force formating the dtime column
        .range(.Cells(FP_end_Row + 1, FP_dTime_Col).Address, .Cells(FP_end_Row + 1 + si_end_Row - si_start_Row, FP_dTime_Col).Address).NumberFormat = "m/d/yyyy h:mm"
    End With
End If

QAwb.Close savechanges:=False

'Part III: combined_QAQC (fd->FDws): find the appending point, then add QAQC data recursively

'get the first QA sheet
For QtrYr_Col = startQA_Col To endQA_Col
    If MyWs.Cells(siteRow, QtrYr_Col).Value <> "" Then
        QAsheet = MyWs.Cells(siteRow, QtrYr_Col).Value
        Exit For
    End If
Next

tgt_start_Row = 2

'loop through each QA sheet, appending data to current QAQC file
Do While QAsheet <> ""
    'open QA_sheet
    If IsWorkBookOpen(QAsheet) = False Then
        Set QAwb = Workbooks.Open(fileName:=QAsheet, UpdateLinks:=False, ReadOnly:=True)
    Else
        Set QAwb = Workbooks(Dir(QAsheet))
    End If
'    With QAwb.Windows(1)
'        .WindowState = xlNormal
'        .Width = 400
'        .Height = 800
'        .Top = 300
'        .Left = 400
'        '.ScrollColumn = 1
'        '.ScrollRow = fd.Cells(14, 1).End(xlDown).Row - 20
'    End With
    
    Set fd = QAwb.Sheets("Flow data")
    'get table defination
    With fd.range("10:20")
        src_dtime_Col = .Find("DateTime").Column
        src_Lvl_Col = .Find("Level 1", lookat:=xlWhole).Column
        src_Flw_Col = .Find("Flow 1", lookat:=xlWhole).Column
        src_Vel_Col = .Find("Vel 1", lookat:=xlWhole).Column
        src_Corr_Flw_Col = .Find("Corrected Flow").Column
        src_Corr_Lvl_Col = .Find("Corrected Level", lookat:=xlWhole).Column
        'find the first row of data, not date
        If fd.Cells(.Find("Level 1").Row + 2, src_Lvl_Col).Value <> "" Then
            src_start_Row = .Find("Level 1").Row + 2
        Else
            src_start_Row = fd.Cells(.Find("Level 1").Row + 2, src_Lvl_Col).End(xlDown).Row
        End If
        'src_start_Row = fd.Columns(src_Lvl_Col).Find("*", fd.Cells(fd.Rows.Count, src_Lvl_Col), xlValues, xlWhole, xlNext).Row
        'src_append_Row = fd.Columns(src_dTime_Col).Find(Format(tgt_append_Date, "mm/dd/yyyy hh:mm:ss"), LookIn:=xlValues).Row
        src_end_Row = fd.Cells(fd.Rows.count, src_Lvl_Col).End(xlUp).Row
    End With
     
    tgt_end_Row = tgt_start_Row + src_end_Row - src_start_Row
     
'    'get the new appending point, determine the next QA sheet to be used
'    tgt_end_Row = FDws.Cells(FDws.Rows.Count, tgt_dTime_Col).End(xlUp).Row
'    tgt_end_Date = FDws.Cells(tgt_end_Row, tgt_dTime_Col).Value
    
    'copy new date
    With FDws
    
'for dates, direct passing formula-calculated value to target cell will cause rounding issue that every midnight is one day backwards
'        .Range(.Cells(tgt_start_Row, tgt_dTime_Col).Address, .Cells(tgt_end_Row, tgt_dTime_Col).Address).Value _
'         = fd.Range(.Cells(src_start_Row, src_dTime_Col).Address, .Cells(src_end_Row, src_dTime_Col).Address).Value

'to bypass this issue, the pastespecial method is used and 'xlPasteValuesAndNumberFormats' argument is selected
        fd.range(.Cells(src_start_Row, src_dtime_Col).Address, .Cells(src_end_Row, src_dtime_Col).Address).Copy
        .range(.Cells(tgt_start_Row, tgt_dTime_Col).Address, .Cells(tgt_end_Row, tgt_dTime_Col).Address).PasteSpecial xlPasteValuesAndNumberFormats
        
        .range(.Cells(tgt_start_Row, tgt_level_Col).Address, .Cells(tgt_end_Row, tgt_level_Col).Address).Value _
         = fd.range(.Cells(src_start_Row, src_Lvl_Col).Address, .Cells(src_end_Row, src_Lvl_Col).Address).Value
        
        .range(.Cells(tgt_start_Row, tgt_velocity_Col).Address, .Cells(tgt_end_Row, tgt_velocity_Col).Address).Value _
         = fd.range(.Cells(src_start_Row, src_Vel_Col).Address, .Cells(src_end_Row, src_Vel_Col).Address).Value
        
        .range(.Cells(tgt_start_Row, tgt_flow_Col).Address, .Cells(tgt_end_Row, tgt_flow_Col).Address).Value _
         = fd.range(.Cells(src_start_Row, src_Flw_Col).Address, .Cells(src_end_Row, src_Flw_Col).Address).Value
        
        .range(.Cells(tgt_start_Row, tgt_corr_flow_Col).Address, .Cells(tgt_end_Row, tgt_corr_flow_Col).Address).Value _
         = fd.range(.Cells(src_start_Row, src_Corr_Flw_Col).Address, .Cells(src_end_Row, src_Corr_Flw_Col).Address).Value
        
        .range(.Cells(tgt_start_Row, tgt_corr_level_Col).Address, .Cells(tgt_end_Row, tgt_corr_level_Col).Address).Value _
         = fd.range(.Cells(src_start_Row, src_Corr_Lvl_Col).Address, .Cells(src_end_Row, src_Corr_Lvl_Col).Address).Value
    
        'force formating the dtime column
        .range(.Cells(tgt_start_Row, tgt_dTime_Col).Address, .Cells(tgt_end_Row, tgt_dTime_Col).Address).NumberFormat = "m/d/yyyy h:mm"
    End With

QAwb.Close savechanges:=False

tgt_start_Row = tgt_end_Row + 1
QtrYr_Col = QtrYr_Col + 1
QAsheet = MyWs.Cells(siteRow, QtrYr_Col).Value
        
'    'get the new appending point, determine the next QA sheet to be used
'    tgt_end_Row = FDws.Cells(FDws.Rows.Count, tgt_dTime_Col).End(xlUp).Row
'    tgt_end_Date = FDws.Cells(tgt_end_Row, tgt_dTime_Col).Value
'    'add a time step for appending point
'    tgt_append_Date = DateAdd("n", intvl, tgt_end_Date)
'    QtrYr2 = "Q" & DatePart("q", tgt_append_Date) & "-" & Right(DatePart("yyyy", tgt_append_Date), 2)
'    'if the new appending point is at a new Quarter sheet, then close the current one and loop to the new sheet
'    'otherwise, exit the loop
'    If QtrYr2 <> QtrYr Then
'        QAwb.Close savechanges:=False
'        QtrYr_Col = MyWs.Rows(1).Find(QtrYr2).Column
'        'get the QA sheet
'        QAsheet = MyWs.Cells(Row, QtrYr_Col).Value
'    Else
'        Exit Do
'    End If
DoEvents
Loop

'update input.csv (step 5)
With Inpwb.Sheets(1)
    .Cells(3, 2).Value = Dir(target1)
    .Cells(2, 2).Value = Dir(target2)
End With

'log operations

Call loginfo(siteName, "Big Picture input files generated (create)")

Application.DisplayAlerts = False
Application.ScreenUpdating = True
'reset sheet numbers in new workbook (global variable)
Application.SheetsInNewWorkbook = 3

End Sub
Private Sub BigPrevertBtn_Click()
'Hao Zhang @ 2015.2.17
'disgard all changes, delete generated folders
'get the destination folder to save generated files


'Dim localpath As String
'localpath = BigPpathTB.Value
''get the row number in CurSitesLB
'Dim Row As Integer
'For isite = 0 To chosenSitesLB.ListCount - 1
'    If chosenSitesLB.Selected(isite, 1) = True Then
'        Row = chosenSitesLB.List(isite, 0)
'        siteName = MyWs.Cells(Row, siteName_Col).Value
'        Exit For
'    End If
'Next

'localpath = BigPpathTB.Value
'For isite = 0 To chosenSitesLB.ListCount - 1
'    If chosenSitesLB.Selected(isite) = True Then
'        siteName = MyWs.Cells(chosenSitesLB.List(isite, 0), siteName_Col).Value
'    End If
'Next

'close all files without saving
'For Each wb In Application.Workbooks
'    If wb.Name Like "*Combined_field_points*" Then
'        'get site name
'        siteName = Left(wb.Name, InStr(1, wb.Name, "_") - 1)
'        wb.Close savechanges:=False
'    End If
'Next

For Each wb In Application.Workbooks
    If wb.Name Like "*Combined_field_points*" Or wb.Name Like "*CombinedQAQC*" Or wb.Name Like "*input*" Or wb.Name Like "*-*(Q?-??)*" Then
        localpath = wb.path
        wbPath = wb.path & "\" & wb.Name
        wb.Close savechanges:=False
        Kill wbPath
    End If
Next

'For Each wb In Application.Workbooks
'    If wb.Name Like "*input*" Then
'        wb.Close savechanges:=False
'    End If
'Next
'
'For Each wb In Application.Workbooks
'    If wb.Name Like "*-*(Q?-??)*" Then
'        wb.Close savechanges:=False
'    End If
'Next

'ask if delete entire folder
'this if is to avoid error that a user may press the button with no worksheet opened
If fso.FolderExists(localpath) = True Then
    If MsgBox("Delete site subfolder as well?", vbYesNo, "Warning") = vbYes Then
        fso.DeleteFolder (localpath)
    End If
End If
''''log operation
'Call loginfo(siteName, "Big Picture input files reverted")

End Sub

Private Sub BigPsaveBtn_Click()
'Hao Zhang @ 2015.2.16
'save updated file

For Each wb In Application.Workbooks
    If wb.Name Like "*Combined_field_points*" Or wb.Name Like "*CombinedQAQC*" Or wb.Name Like "*input*" Or wb.Name Like "*-*(Q?-??)*" Then
'        'get site name
'        siteName = Left(wb.Name, InStr(1, wb.Name, "_") - 1)
        wb.Close savechanges:=True
    End If
Next

'For Each wb In Application.Workbooks
'    If wb.Name Like "*CombinedQAQC*" Then
'        wb.Close savechanges:=True
'    End If
'Next
'
'For Each wb In Application.Workbooks
'    If wb.Name Like "*input*" Then
'        wb.Close savechanges:=True
'    End If
'Next
'
'For Each wb In Application.Workbooks
'    If wb.Name Like "*-*(Q?-??)*" Then
'        wb.Close savechanges:=False
'    End If
'Next

''''log operation
'Call loginfo(siteName, "Big Picture input files saved")

''move to the next item in the listbox
''this step must be on bottom
''(don't move sites down yet)
'Call MoveDoneSiteBtn_Click

End Sub

Private Sub BigPrunRbtn_Click()
'Hao Zhang @ 2015.2.22
'Run one or a series of Run_BigPicture.bat of selected sites
'''this sub could not address the issue when there is an error in batch operation
Dim objShell As Object
Dim objWshScriptExec As Object
Dim objStdOut As Object
Dim rline As String
Dim strline As String
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1

Set objShell = CreateObject("WScript.Shell")
'copy Big_Picture.R to temp folder
fso.CopyFile "\\pwdhqr\oows\Modeling\Data\Temporary Monitors\Flow Monitoring\QAQC Procedures_CDM\BigPicture\BigPicture_v1.5.R", fso.GetSpecialFolder(2) & "\BigPicture.R", True

For isite = 0 To chosenSitesLB.ListCount - 1
    If chosenSitesLB.Selected(isite) = True Then
        siteName = chosenSitesLB.List(isite, 1)
        BigPPath = BigPpathTB.Value & "\" & siteName
        
        'delete previous output in big_picture folder
        If fso.FolderExists(BigPPath & "\big_picture") = True Then
            fso.DeleteFolder (BigPPath & "\big_picture")
        End If
        
        'command-line statements to execute Big_Picture.R
        'pushd works like cd, but it allows to change disk while cd can't
        '&& to connect two statements, when the first succeed, then exectute the second
        'Rscript.exe must be added to the environment variables list, otherwise, full path is required
        BPcmd = "pushd " & BigPPath & " && Rscript.exe %TEMP%\BigPicture.R && exit"
        '/C is a parameter that closes the window when program terminates
        '/K : run the command then return to the CMD prompt
        'ErrCode = objShell.Run("cmd.exe /c " & BPcmd, Style, waitOnReturn)
        Set objWshScriptExec = objShell.Exec("cmd.exe /c " & BPcmd)
        Set objStdOut = objWshScriptExec.StdOut
        strline = objStdOut.ReadAll
        If InStr(1, strline, "Printing plots... Finished.") <> 0 Then
            Call loginfo(siteName, "Big Picture run successful")
        Else
            Call loginfo(siteName, "Big Picture run failed")
        End If
    End If
Next


''create a temp bat file that runs selected sites
'tempBAT = BigPpathTB.Value & "\tempBAT.bat"
'ifileNum = FreeFile
''store the tempBAT in the temp folder
''tempBAT = fso.GetSpecialFolder(2) & "\" & ifileNum & ".bat"
'Open tempBAT For Output As #ifileNum
'        Print #ifileNum, "Set scriptpath=" & Chr(34) & "\\pwdhqr\oows\Modeling\Data\Temporary Monitors\Flow Monitoring\QAQC Procedures_CDM\BigPicture\BigPicture_v1.5.R" & Chr(34)
'        Print #ifileNum, "set temppath=%TEMP%\Big_Picture.R"
'        Print #ifileNum, "copy %scriptpath% %temppath%"
'        Print #ifileNum, "Set rpath =" & Chr(34) & "C:\Program Files\R\R-3.0.2\bin\Rscript.exe" & Chr(34)
'
'For isite = chosenSitesLB.ListCount - 1 To 0 Step -1
'    If chosenSitesLB.Selected(isite) = True Then
'        siteName = chosenSitesLB.List(isite, 1)
'        BigPPath = BigPpathTB.Value & "\" & siteName
'        Print #ifileNum, "set calldir=" & Chr(34) & BigPPath & Chr(34)
'        Print #ifileNum, "pushd %calldir%"
'        Print #ifileNum, "%rpath% %temppath%"
'    End If
'Next
'
'    Close #ifileNum
'ErrorCode = objShell.Run(tempBAT, Style, waitOnReturn)
    'Shell BigPbatPath, vbNormalFocus

'Kill tempBAT
'For isite = chosenSitesLB.ListCount - 1 To 0 Step -1
'    If chosenSitesLB.Selected(isite) = True Then
'        siteName = chosenSitesLB.List(isite, 1)
'        BigPbatPath = BigPpathTB.Value & "\" & siteName & "\Run_BigPicture.bat"
'       'wscript is better than shell because it has waitOnReturn option, which let bat file run sequentially
'        objShell.Run BigPbatPath, windowStyle, waitOnReturn
'       'get the shell output (to check whether the bat file is executed successfully)
'        Set objWshScriptExec = objShell.Exec(BigPbatPath)
'        Set objStdOut = objWshScriptExec.StdOut
'        strline = objStdOut.ReadAll
'       'get output line by line (abandoned)
'        While Not objStdOut.AtEndOfStream
'            rline = objStdOut.ReadLine
'            If rline <> "" Then strline = strline & vbCrLf & rline
'       'you can handle the results as they are written to and subsequently read from the StdOut object
'        Wend
'        If InStr(1, strline, "Printing plots... Finished.") <> 0 Then
'            Call loginfo(siteName, "Run_BigPicture.bat executed successfully.")
'        Else
'            Call loginfo(siteName, "Run_BigPicture.bat failed.")
'        End If
'        'automate the 'press any key' process
'        'SendKeys "a"
'        Set objWshScriptExec = Nothing
'        Set objStdOut = Nothing
'        strline = ""
'        Call MoveDoneSiteBtn_Click
'        'move processed sites down
'        DoneSitesLB.AddItem (chosenSitesLB.List(isite))
'        DoneSitesLB.List(DoneSitesLB.ListCount - 1, 1) = chosenSitesLB.List(isite, 1)
'
'        decide single/burst operation
'        If BigPChtBurstOB = False Then
'            Exit Sub
'        End If
'    End If
'    DoEvents
'Next

''removed chosen sites after loop ends. otherwise, the ListCount will altered for each loop which leads to error
'For isite = chosenSitesLB.ListCount - 1 To 0
'    If chosenSitesLB.Selected(isite) = True Then
'        chosenSitesLB.RemoveItem (isite)
'End If
'Next

End Sub
Private Sub BigPupdFlBtn_Click()
'Hao Zhang @ 2015.2.17, rev @ 2015.3.13
'move the FD and FP files to bk folder on server
'move generated big_picture back to server, overwrite existing version

'Initialization
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'get the destination folder to save generated files
Dim localpath As String
localpath = BigPpathTB.Value
'get the row number in CurSitesLB
For isite = 0 To chosenSitesLB.ListCount - 1
    If chosenSitesLB.Selected(isite) = True Then
        siteRow = chosenSitesLB.List(isite, 0)
        siteName = chosenSitesLB.List(isite, 1)
    
    'make sure the wildcard is used so all files within the site folder will be affected
    srcfldr = localpath & "\" & siteName & "\*"

    'get target file folder, dealing with naming variations
    Dim tgtfldr As String
    If fso.FolderExists(MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\BigPicture\") = True Then
        tgtfldr = MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\BigPicture\"
    ElseIf fso.FolderExists(MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\Big Picture\") = True Then
        tgtfldr = MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\Big Picture\"
    Else
        MsgBox "please specify the Big Picture folder path on server:"
        tgtfldr = GetFldr(MyWs.Cells(siteRow, siteFldr_Col).Value & "\QAQC\") & "\"
    End If
    
    
    QA_old = Dir(tgtfldr & "*QAQC*.csv")
    fp_old = Dir(tgtfldr & "*field_points*.csv")
    QA_new = Dir(srcfldr & "*QAQC*.csv")
    fp_new = Dir(srcfldr & "*field_points*.csv")
    'seems the folder name is not case-sensitive
    tgt_bk = tgtfldr & "bk\"
    
    'create a bk folder if not exist
    If fso.FolderExists(tgt_bk) = False Then
        fso.CreateFolder (tgt_bk)
    End If
    
    'operation on QA file
    If QA_old <> "" Then
        If fso.FileExists(tgt_bk & QA_old) = True Then
            'cond1: QA file exist in both tgtfldr and bk: delete the one in tgtfldr
            Kill tgtfldr & QA_old
        Else
            'cond2: QA file exist in tgtfldr, but not bk: move existing file to bk
            Name tgtfldr & QA_old As tgt_bk & QA_old
        End If
        'cond3: QA file not exist in tgtfldr: do nothing
    End If
    
    'operation on fp file
    If fp_old <> "" Then
        If fso.FileExists(tgt_bk & fp_old) = True Then
            'cond1: fp file exist in both tgtfldr and bk: delete the one in tgtfldr
            Kill tgtfldr & fp_old
        Else
            'cond2: fp file exist in tgtfldr, but not bk: move existing file to bk
            Name tgtfldr & fp_old As tgt_bk & fp_old
        End If
        'cond3: fp file not exist in tgtfldr: do nothing
    End If
    
    'overwrites folder and files on server
    fso.CopyFolder srcfldr, tgtfldr, True
    fso.CopyFile srcfldr, tgtfldr, True
    
    Call loginfo(siteName, "Big Picture files uploaded to server")
    End If
DoEvents
Next

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub BigPChtBurstOB_Click()
'Hao Zhang @ 2015.2.22
'when Burst mode is clicked, pre-select all chosen sites
'user can still pick sites to be processed
For isite = 0 To chosenSitesLB.ListCount - 1
    chosenSitesLB.Selected(isite) = True
Next
End Sub

Private Sub BigPpptBrowseBtn_Click()
'Hao Zhang @ 2015.3.17
'browse to output folder
'allow user selecting a folder to save the Big Picture Input files

'use this variable to save user's selection
Dim tempPath As String
tempPath = GetFldr(sPath)
'if no path was specified, then retain the original value
If tempPath <> "" Then
    BigPpathTB.Value = tempPath
End If

End Sub
Private Sub BigPQApptBtn_Click()

'Hao Zhang @ 2015.3.17
'Auto-generates the 2 ppt files that contain monthly TS charts based on the yellow sheet
'need to add "Microsoft PowerPoint 14.0 Object library" to the references first

Application.ScreenUpdating = False

'PART 0: retrieve yellow sheet
'specify the month and year
Dim iMon As Integer
Dim iYear As Integer
iMon = BigPpptMonCB.Text
iYear = BigPpptYrCB.Text

Dim ylwWb As Workbook
If ylwsheet = "" Then
    If MsgBox("Please select the yellow sheet you want to refer to:", vbOKCancel) = vbOK Then
        ylwPath = "M:\Data\Temporary Monitors\Flow Monitoring\Supplementary Documents\CSL Meeting Notes\" & iYear & "\" & MonthName(iMon + 1) & " " & iYear & "\"
        ylwsheet = GetFile(ylwPath)
    Else
        'cancel the operation
        Exit Sub
    End If
End If

If ylwsheet <> "" Then
    If IsWorkBookOpen(ylwsheet) = False Then
        Set ylwWb = Workbooks.Open(fileName:=ylwsheet)
    Else
        Set ylwWb = Workbooks(Dir(ylwsheet))
    End If
Else
    MsgBox "No file is selected."
    Exit Sub
End If

'get table defination of the yellow sheet
With ylwWb.Sheets(1)
    Comments_Col = .range("1:5").Find("Comments").Column
    RDII_Row = .Columns(1).Find("OOW RDII Monitoring Sites:").Row
    DCIA_Row = .Columns(1).Find("OOW DCIA Monitoring Sites:").Row
    SWM_Row = .Columns(1).Find("Stormwater Monitoring Sites:").Row
End With

'PART I: create good/bad QAQC ppt
Dim newppt As New PowerPoint.Application
Dim CurSlide As PowerPoint.Slide

Dim goodPPT As Presentation
Set goodPPT = newppt.Presentations.Add(msoCTrue)
pptName = GetUniqueName(BigPpptPathTB.Text & "\" & MonthName(iMon) & "_QAQC_Good.pptx")
goodPPT.SaveAs pptName

Dim badPPT As Presentation
Set badPPT = newppt.Presentations.Add(msoCTrue)
pptName = GetUniqueName(BigPpptPathTB.Text & "\" & MonthName(iMon) & "_QAQC_Bad.pptx")
badPPT.SaveAs pptName

'Add the title slide in the good PowerPoint
Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutTitleOnly)
    'adjust the textbox layout
    With CurSlide.Shapes(1)
        .TextFrame.TextRange.Text = MonthName(iMon) & " " & iYear & " Data" & Chr(13) & "QAQC Good Sites"
        .Width = 620
        .Height = 200
        .Left = 50
        .Top = 100
    End With
'Add the cover slide in the bad PowerPoint
Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutTitleOnly)
'adjust the textbox layout
    With CurSlide.Shapes(1)
        .TextFrame.TextRange.Text = MonthName(iMon) & " " & iYear & " Data" & Chr(13) & "QAQC Bad Sites"
        .Width = 620
        .Height = 200
        .Left = 50
        .Top = 100
    End With
'set the window to just show slides
'newPPT.ActiveWindow.ViewType = ppViewSlide

'PART II: add plots to QAQC ppt
With ylwWb.Sheets(1)
    For iRow = RDII_Row To SWM_Row + 2
        'Add plots
        'don't use two '<' for ranges, as VBA will consider them with a 'or' operand thus includes all conditions
        If (RDII_Row < iRow And iRow < DCIA_Row - 2) Or (DCIA_Row < iRow And iRow < SWM_Row - 2) Or (iRow > SWM_Row) Then
            'get QA sheet
            siteName = .Cells(iRow, 1).Value
            siteRow = MyWs.Columns(siteName_Col).Find(siteName).Row
            QtrYr = "Q" & DatePart("q", DateSerial(iYear, iMon, 1)) & "-" & Right(iYear, 2)
            siteCol = MyWs.Rows(1).Find(QtrYr).Column
            QAsheet = MyWs.Cells(siteRow, siteCol).Value
            If fso.FileExists(QAsheet) = False Then
                'skip the site if no QA sheet could be found
                loginfo siteName, "Cannot find QAQC file on server (skipped)"
                GoTo NextIteration
            Else
                If IsWorkBookOpen(QAsheet) = False Then
                    Set QAwb = Workbooks.Open(fileName:=QAsheet, ReadOnly:=True, UpdateLinks:=False)
                Else
                    Set QAwb = Workbooks(Dir(QAsheet))
                End If
                'hide the QA sheet
                QAwb.Windows(1).visible = False
            End If
            
            'determine which ppt to paste to
            If .Cells(iRow, Comments_Col).Font.Color = vbRed Then
                Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutBlank)
            Else
                Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutBlank)
            End If
            'copy TS plot to ppt
            QAwb.Charts(MonthName(iMon, True) & " TS").ChartArea.Copy
            CurSlide.Shapes.PasteSpecial DataType:=ppPasteDefault
            With CurSlide.Shapes(1)
                .Width = 680  '680
                .Height = 520 '520
                .Left = 20    '20
                .Top = 10
            End With
            
            'determine which ppt to paste to
            If .Cells(iRow, Comments_Col).Font.Color = vbRed Then
                Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutBlank)
            Else
                Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutBlank)
            End If
            'copy TS CORR plot to ppt
            QAwb.Charts(MonthName(iMon, True) & " TS CORR").ChartArea.Copy
            CurSlide.Shapes.PasteSpecial DataType:=ppPasteDefault
            With CurSlide.Shapes(1)
                .Width = 680  '680
                .Height = 520 '520
                .Left = 20    '20
                .Top = 10
            End With
            QAwb.Close savechanges:=False
            
        'add RDII cover page to both ppt
        ElseIf iRow = RDII_Row Then
            'add RDII cover pages for good & bad sites
            'good
            Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutTitleOnly)
            'adjust the textbox layout
            With CurSlide.Shapes(1)
                .TextFrame.TextRange.Text = "RDII sites"
                .Width = 620
                .Height = 200
                .Left = 50
                .Top = 100
            End With
            'bad
            Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutTitleOnly)
            'adjust the textbox layout
            With CurSlide.Shapes(1)
                .TextFrame.TextRange.Text = "RDII sites"
                .Width = 620
                .Height = 200
                .Left = 50
                .Top = 100
            End With
            
        'add DCIA cover page to both ppt
        ElseIf iRow = DCIA_Row Then
            'add DCIA cover pages for good & bad sites
            'good
            Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutTitleOnly)
            'adjust the textbox layout
            With CurSlide.Shapes(1)
                .TextFrame.TextRange.Text = "DCIA sites"
                .Width = 620
                .Height = 200
                .Left = 50
                .Top = 100
            End With
            'bad
            Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutTitleOnly)
            'adjust the textbox layout
            With CurSlide.Shapes(1)
                .TextFrame.TextRange.Text = "DCIA sites"
                .Width = 620
                .Height = 200
                .Left = 50
                .Top = 100
            End With
        End If
    DoEvents
NextIteration:
    Next
End With
loginfo "all sites", "QAQC ppt generated"
Application.ScreenUpdating = True

'thought: could create a collection to simplify the code

End Sub

Private Sub BigPbpPPTbtn_Click()
'Hao Zhang @ 2015.3.18
'create two ppts for Big Picture TS plots for good/bad sites, based on monthly yellow sheet
'need to add "Microsoft PowerPoint 14.0 Object library" to the references first

Application.ScreenUpdating = False

'PART 0: retrieve yellow sheet
'specify the month and year
Dim iMon As Integer
Dim iYear As Integer
iMon = BigPpptMonCB.Text
iYear = BigPpptYrCB.Text

'ask the user to specify the yellow sheet
'use a global variable for the path of ylwsheet, so user only need to specify it once

Dim ylwWb As Workbook
If ylwsheet = "" Then
    If MsgBox("Please select the yellow sheet you want to refer to:", vbOKCancel) = vbOK Then
        ylwPath = "M:\Data\Temporary Monitors\Flow Monitoring\Supplementary Documents\CSL Meeting Notes\" & iYear & "\" & MonthName(iMon + 1) & " " & iYear & "\"
        ylwsheet = GetFile(ylwPath)
    Else
        'cancel the operation
        Exit Sub
    End If
End If

If ylwsheet <> "" Then
    If IsWorkBookOpen(ylwsheet) = False Then
        Set ylwWb = Workbooks.Open(fileName:=ylwsheet)
    Else
        Set ylwWb = Workbooks(Dir(ylwsheet))
    End If
Else
    MsgBox "No file is selected."
    Exit Sub
End If

'get table defination of the yellow sheet
With ylwWb.Sheets(1)
    Comments_Col = .range("1:5").Find("Comments").Column
    RDII_Row = .Columns(1).Find("OOW RDII Monitoring Sites:").Row
    DCIA_Row = .Columns(1).Find("OOW DCIA Monitoring Sites:").Row
    SWM_Row = .Columns(1).Find("Stormwater Monitoring Sites:").Row
End With

'PART I: create good/bad Big Picture ppt
Dim newppt As New PowerPoint.Application
Dim CurSlide As PowerPoint.Slide

Dim goodPPT As Presentation
Set goodPPT = newppt.Presentations.Add(msoCTrue)
pptName = GetUniqueName(BigPpptPathTB.Text & "\" & MonthName(iMon) & "_BigPicture_Good.pptx")
goodPPT.SaveAs pptName

Dim badPPT As Presentation
Set badPPT = newppt.Presentations.Add(msoCTrue)
pptName = GetUniqueName(BigPpptPathTB.Text & "\" & MonthName(iMon) & "_BigPicture_Bad.pptx")
badPPT.SaveAs pptName

'Add the cover slide in the good PowerPoint
Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutTitleOnly)
'adjust the textbox layout
With CurSlide.Shapes(1)
    .TextFrame.TextRange.Text = MonthName(iMon) & " " & iYear & " Data" & Chr(13) & "Big Picture Good Sites"
    .Width = 620
    .Height = 200
    .Left = 50
    .Top = 100
End With
'Add the cover slide in the bad PowerPoint
Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutTitleOnly)
'adjust the textbox layout
With CurSlide.Shapes(1)
    .TextFrame.TextRange.Text = MonthName(iMon) & " " & iYear & " Data" & Chr(13) & "Big Picture Good Sites"
    .Width = 620
    .Height = 200
    .Left = 50
    .Top = 100
End With

'set the window to just show slides
'newPPT.ActiveWindow.ViewType = ppViewSlide

'PART II: add plots to Big Picture ppt


'add RDII plots to powerpoint
With ylwWb.Sheets(1)
    For iRow = RDII_Row To SWM_Row + 2
        'Add plots
        If (RDII_Row < iRow And iRow < DCIA_Row - 2) Or (DCIA_Row < iRow And iRow < SWM_Row - 2) Or (iRow > SWM_Row) Then
            'decide which ppt to paste to
            If .Cells(iRow, Comments_Col).Font.Color = vbRed Then
                Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutTitleOnly)
            Else
                Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutTitleOnly)
            End If
            'get src png path
            siteName = .Cells(iRow, 1).Value
            siteFldr = MyWs.Columns(siteFldr_Col).Find(siteName).Value
            'get BigPicture folder path, dealing with naming variations
            Dim BPfldr As String
            If fso.FolderExists(siteFldr & "\QAQC\BigPicture\big_picture\") = True Then
                BPfldr = siteFldr & "\QAQC\BigPicture\big_picture\"
            ElseIf fso.FolderExists(siteFldr & "\QAQC\Big Picture\big_picture\") = True Then
                BPfldr = siteFldr & "\QAQC\Big Picture\big_picture\"
            Else
                loginfo siteName, "Cannot find BigPicture folder on server (skipped)"
                'skip the current site and move on
                GoTo NextIteration
            End If
            TSpng = BPfldr & "uncorrected_ts_(All).png"
            TSCORRpng = BPfldr & "corrected_ts_(All).png"
            With CurSlide.Shapes(1)
                .TextFrame.TextRange.Text = siteName
                .Width = 620
                .Height = 50
                .Left = 50
                .Top = 20
            End With
            'paste pics to powerpoint
            'the parameters of shapes.AddPicture are:
            'path: pic path
            'linktofile:(bol) refering to the pic only (not phyiscally copy to the ppt)?
            'savetodocument: (bol) save the pic in the ppt?
            'left, top, width, height: position and dimension, use -1 in dim to display the "natural" look (looks like fill the slide)
            CurSlide.Shapes.AddPicture TSpng, msoFalse, msoTrue, 0, 70, -1, -1
            
            'decide which ppt to paste to
            If .Cells(iRow, Comments_Col).Font.Color = vbRed Then
                Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutTitleOnly)
            Else
                Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutTitleOnly)
            End If
            With CurSlide.Shapes(1)
                .TextFrame.TextRange.Text = siteName
                .Width = 620
                .Height = 50
                .Left = 50
                .Top = 20
            End With
            CurSlide.Shapes.AddPicture TSCORRpng, msoFalse, msoTrue, 0, 70, -1, -1
    
        'add RDII cover page to both ppt
        ElseIf iRow = RDII_Row Then
            'add RDII sites cover pages for good & bad sites
            'good
            Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutTitleOnly)
                'adjust the textbox layout
                With CurSlide.Shapes(1)
                    .TextFrame.TextRange.Text = "RDII sites"
                    .Width = 620
                    .Height = 200
                    .Left = 50
                    .Top = 100
                End With
            'bad
            Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutTitleOnly)
                'adjust the textbox layout
                With CurSlide.Shapes(1)
                    .TextFrame.TextRange.Text = "RDII sites"
                    .Width = 620
                    .Height = 200
                    .Left = 50
                    .Top = 100
                End With
            
        'add DCIA cover page to both ppt
        ElseIf iRow = DCIA_Row Then
            'Add the cover page slide in PowerPoint (DCIA)
            Set CurSlide = goodPPT.Slides.Add(goodPPT.Slides.count + 1, ppLayoutTitleOnly)
            'adjust the textbox layout
            With CurSlide.Shapes(1)
                .TextFrame.TextRange.Text = "DCIA sites"
                .Width = 620
                .Height = 200
                .Left = 50
                .Top = 100
            End With
            'Add the cover page slide in PowerPoint (DCIA)
            Set CurSlide = badPPT.Slides.Add(badPPT.Slides.count + 1, ppLayoutTitleOnly)
            With CurSlide.Shapes(1)
                .TextFrame.TextRange.Text = "DCIA sites"
                .Width = 620
                .Height = 200
                .Left = 50
                .Top = 100
            End With
        End If
    DoEvents
NextIteration:
    Next
End With
loginfo "all sites", "BigPicture ppt generated"

Application.ScreenUpdating = True

End Sub

'****************************************************************
'*************************SSOAP**********************************
'****************************************************************


Private Sub SSOAPbrowseBtn_Click()
'Hao Zhang @ 2015.3.13
'allow user selecting a folder to save the SSOAP Input files

'use this variable to save user's selection
tempPath = GetFldr(sPath)
'if no path was specified, then retain the original value
If tempPath <> "" Then
    SSOAPpathTB.Value = tempPath
End If

End Sub


Private Sub SSOAPfldrCreateBtn_Click()
'Hao Zhang @ 2015.3.13
'create SSOAP folder inside the specified path

If fso.FolderExists(SSOAPpathTB.Value & "\SSOAP") = False Then
    SSOAPpathTB.Value = fso.CreateFolder(SSOAPpathTB.Value & "\SSOAP").path
Else
    MsgBox "The specified folder already exists."
    SSOAPpathTB.Value = SSOAPpathTB.Value & "\SSOAP"
End If

End Sub

Private Sub SSOAPinputGenBtn_Click()
'Hao Zhang @ 2015.3.13
'generate SSOAP input files (flow, rain)
Dim isite As Integer

'check if chosenSiteLB is not empty and at least one site was selected
If chosenSitesLB.ListCount > 0 And chosenSitesLB.ListIndex <> -1 Then
    For isite = 0 To chosenSitesLB.ListCount - 1
        If chosenSitesLB.Selected(isite) = True Then
            siteRow = chosenSitesLB.List(isite, 0)
            siteName = MyWs.Cells(siteRow, siteName_Col).Value
            
            If SSOAPappOB = True Then
                Call SSOAPappend(siteName)
            Else
                Call SSOAPcomplete(siteName)
            End If
            
            Call loginfo(siteName, "SSOAP input files generated")
         End If
         
         If chosenSitesLB.ListIndex = 0 Then
            'single mode
            Exit Sub
         ElseIf chosenSitesLB.ListIndex > 0 Then
            'burst mode
            Call SSOAPsaveBtn_Click
         End If
    DoEvents
    Next
Else
    MsgBox "please add sites in 'Chosen Sites' first", vbCritical, "Error"
    Exit Sub
End If
End Sub

Sub SSOAPappend(siteName As String)
'Hao Zhang @ 2015.3.13
'1. copy the most recent Rainfall and FLOWINPUT files to local drive
'2. find the end point, append new data from there

'PART 0 : Initialization
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'get the destination folder to save generated files
Dim localpath As String
localpath = SSOAPpathTB.Value
'get the row number in CurSitesLB
siteRow = MyWs.Columns(siteName_Col).Find(siteName).Row

'get source file folder, dealing with naming variations
Dim srcfldr As String

'create SSOAP folder if it's not exist
If fso.FolderExists(MyWs.Cells(siteRow, siteFldr_Col).Value & "\SSOAP") = False Then
    MsgBox "there is no SSOAP files to append, please use the 'complete' method."
    Exit Sub
Else
    'find the latest SSOAP folder
    srcfldr = MyWs.Cells(siteRow, siteFldr_Col).Value & "\SSOAP\"
    For Each Fldr1 In fso.GetFolder(srcfldr).SubFolders
        For Each Fldr2 In fso.GetFolder(srcfldr).SubFolders
            If Fldr1.DateLastModified >= Fldr2.DateLastModified Then
                srcfldr = Fldr1.path
            Else
                srcfldr = Fldr2.path
            End If
        Next
    Next
End If

Dim inpfldr As String
inpfldr = srcfldr & "\"
srcfldr = fso.GetParentFolderName(srcfldr) & "\"
            
'PART I :file handling

'find the first file that partially matches the source file name
'only the filename with extension is assigned to sFound1 and sFound2
' * is used as a wildcard

'set src path of rainfall files
Dim sFound1 As String
Dim source1 As String
sFound1 = Dir(inpfldr & "*_Rainfall_*.csv")
If sFound1 <> "" Then
    source1 = inpfldr & sFound1
Else
    MsgBox "Please specify 'Rainfall' file:"
    source1 = GetFile(inpfldr)
End If

'set src path of flow files
Dim sFound2 As String
Dim source2 As String
sFound2 = Dir(inpfldr & "*_FLOWINPUT_*.csv")
If sFound2 <> "" Then
    source2 = inpfldr & sFound2
Else
    MsgBox "Please specify 'Flow' file:"
    source2 = GetFile(inpfldr)
End If

'set src path of SSOAP file
Dim sFound3 As String
Dim source3 As String
sFound3 = Dir(srcfldr & "*SSOAP*.sdb")
If sFound3 <> "" Then
    source3 = srcfldr & sFound3
Else
    MsgBox "Please specify 'SSOAP.sdb' file:"
    source3 = GetFile(srcfldr)
End If

If IsNull(SSOAPiniTB.Text) = False Then
    ini = SSOAPiniTB.Text
Else
    ini = InputBox("please enter your initial.")
End If

'set the full path of target files
Dim target1 As String
Dim target2 As String
Dim target3 As String
tgtfldr = localpath & "\" & siteName & "\" & "SSOAP_" & Format(Now(), "yymmdd") & "_" & ini & "\"
target1 = tgtfldr & "Input\" & siteName & "_Rainfall_Radar2Sheds_" & Format(Now(), "yymmdd") & "_" & ini & ".csv"
target2 = tgtfldr & "Input\" & siteName & "_SSOAP_FLOWINPUT_" & Format(Now(), "yymmdd") & "_" & ini & ".csv"
target3 = tgtfldr & siteName & "_SSOAP_" & Format(Now(), "yymmdd") & "_" & ini & ".sdb"

'if site subfolder is not exist, create the folder first
If fso.FolderExists(localpath & "\" & siteName & "\") = False Then
    fso.CreateFolder localpath & "\" & siteName & "\"
End If

'if site subfolder exists but the folder for current SSOAP run is not exist
If fso.FolderExists(tgtfldr) = False Then
    fso.CreateFolder tgtfldr
Else
    'clear all contents in the tgtfldr
    fso.DeleteFile tgtfldr & "*"
End If

If fso.FolderExists(tgtfldr & "input") = False Then
    fso.CreateFolder tgtfldr & "input"
Else
    'clear all contents in the tgtfldr
    fso.DeleteFile tgtfldr & "input\*"
End If

'copy/overwrite rainfall from server to local drive
fso.CopyFile source1, target1
'copy/overwrite flow from server to local drive
fso.CopyFile source2, target2
'copy/overwrite SSOAP from server to local drive
fso.CopyFile source3, target3

'PART II : append flow

'open Fl file on local drive
Dim Flwb As Workbook
If IsWorkBookOpen(target2) = False Then
    Set Flwb = Workbooks.Open(fileName:=target2)
Else
    Set Flwb = Workbooks(Dir(target2))
End If

'adjust windows for visual check
With Flwb.Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth * 0.5
    .Height = Application.UsableHeight * 0.6
    .Top = 0
    .Left = Application.UsableWidth * 0.5
    .ScrollColumn = 1
    .ScrollRow = 1
    ''scroll to the end
    '.ScrollRow = FDwb.Sheets(1).Cells(1, 1).End(xlDown).Row - 10
End With

Dim Flws As Worksheet
Set Flws = Flwb.Sheets(1)
'get table defination
With Flws.Rows(1)
    Fl_Month_Col = .Find("Month", MatchCase:=False).Column
    Fl_Day_Col = .Find("Day", MatchCase:=False).Column
    Fl_Year_Col = .Find("Year", MatchCase:=False).Column
    Fl_Hour_Col = .Find("Hour", MatchCase:=False).Column
    Fl_Minute_Col = .Find("Minute", MatchCase:=False).Column
    Fl_flow_Col = .Find("Flow", MatchCase:=False).Column
End With

'get the appending point, determine the first QA sheet to be used
Dim tgt_end_Row As Long
Dim tgt_append_Date As Date

'so far SSOAP are only based on 15 min data, but it can be refined
intvl = 15
tgt_end_Row = Flws.Cells(Flws.Rows.count, Fl_Month_Col).End(xlUp).Row
tgt_end_Date = DateSerial(Flws.Cells(tgt_end_Row, Fl_Year_Col).Value, Flws.Cells(tgt_end_Row, Fl_Month_Col).Value, Flws.Cells(tgt_end_Row, Fl_Day_Col).Value) + TimeSerial(Flws.Cells(tgt_end_Row, Fl_Hour_Col).Value, Flws.Cells(tgt_end_Row, Fl_Minute_Col).Value, 0)

'add a time step for appending point
'DateAdd(): "m"=month, "n"=minute,"yyyy"=year, "y"=day of the year, "w"=weekday, "ww"=week, "q"=quarter
tgt_append_Date = DateAdd("n", intvl, tgt_end_Date)
QtrYr = "Q" & DatePart("q", tgt_append_Date) & "-" & Right(DatePart("yyyy", tgt_append_Date), 2)
QtrYr_Col = MyWs.Rows(1).Find(QtrYr).Column

'get the lastest QA sheet
QAsheet = MyWs.Cells(siteRow, QtrYr_Col).Value

'loop through each QA sheet, appending data to Flwb
Do While QAsheet <> ""
    'open QA_sheet
    If IsWorkBookOpen(QAsheet) = False Then
        Set QAwb = Workbooks.Open(fileName:=QAsheet, UpdateLinks:=False, ReadOnly:=True)
    Else
        Set QAwb = Workbooks(fso.GetFileName(QAsheet))
    End If
    
    Set fd = QAwb.Sheets("Flow data")
    'get table defination
    With fd.range("10:20")
        src_dtime_Col = .Find("DateTime").Column
        src_Lvl_Col = .Find("Level 1").Column
        src_Corr_Flw_Col = .Find("Corrected Flow").Column
        'force the format of datetime column
        fd.Columns(src_dtime_Col).NumberFormat = "m/d/yyyy h:mm:ss"
        src_append_Row = fd.Columns(src_dtime_Col).Find(Format(tgt_append_Date, "m/d/yyyy h:mm:ss"), LookIn:=xlValues).Row
        src_end_Row = fd.Cells(fd.Rows.count, src_Lvl_Col).End(xlUp).Row
    End With
    
    'put the date in an empty column (will be deleted later)
    Fl_dTime_Col = 7
    
    'copy new date
    If src_end_Row > src_append_Row Then
        With Flws
            'copy dTime to an empty column in tgt
            fd.range(.Cells(src_append_Row, src_dtime_Col).Address, .Cells(src_end_Row, src_dtime_Col).Address).Copy
            .range(.Cells(tgt_end_Row + 1, Fl_dTime_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_dTime_Col).Address).PasteSpecial xlPasteValuesAndNumberFormats
            'stripping dTime into components
            .range(.Cells(tgt_end_Row + 1, Fl_Month_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_Month_Col).Address).Formula = "=month(" & .Cells(tgt_end_Row + 1, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            .range(.Cells(tgt_end_Row + 1, Fl_Day_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_Day_Col).Address).Formula = "=day(" & .Cells(tgt_end_Row + 1, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            .range(.Cells(tgt_end_Row + 1, Fl_Year_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_Year_Col).Address).Formula = "=year(" & .Cells(tgt_end_Row + 1, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            .range(.Cells(tgt_end_Row + 1, Fl_Hour_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_Hour_Col).Address).Formula = "=hour(" & .Cells(tgt_end_Row + 1, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            .range(.Cells(tgt_end_Row + 1, Fl_Minute_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_Minute_Col).Address).Formula = "=minute(" & .Cells(tgt_end_Row + 1, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            'remove formula by pasting the value
            tmpRng = .range(.Cells(tgt_end_Row + 1, Fl_Month_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_Minute_Col).Address)
            .range(.Cells(tgt_end_Row + 1, Fl_Month_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_Minute_Col).Address).Value = tmpRng
            'remove dTime column
            .range(.Cells(tgt_end_Row + 1, Fl_dTime_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_dTime_Col).Address).Clear
            'paste flow data
            .range(.Cells(tgt_end_Row + 1, Fl_flow_Col).Address, .Cells(tgt_end_Row + 1 + src_end_Row - src_append_Row, Fl_flow_Col).Address).Value _
             = fd.range(.Cells(src_append_Row, src_Corr_Flw_Col).Address, .Cells(src_end_Row, src_Corr_Flw_Col).Address).Value
        End With
    Else
        'break condition: when no more data could be added, exit the procedure
        Exit Do
    End If
    
    
    'get the new appending point, determine the next QA sheet to be used
    tgt_end_Row = Flws.Cells(Flws.Rows.count, Fl_Month_Col).End(xlUp).Row
    tgt_end_Date = DateSerial(Flws.Cells(tgt_end_Row, Fl_Year_Col).Value, Flws.Cells(tgt_end_Row, Fl_Month_Col).Value, Flws.Cells(tgt_end_Row, Fl_Day_Col).Value) + TimeSerial(Flws.Cells(tgt_end_Row, Fl_Hour_Col).Value, Flws.Cells(tgt_end_Row, Fl_Minute_Col).Value, 0)
    'add a time step for appending point
    'DateAdd(): "m"=month, "n"=minute,"yyyy"=year, "y"=day of the year, "w"=weekday, "ww"=week, "q"=quarter
    tgt_append_Date = DateAdd("n", intvl, tgt_end_Date)
    QtrYr2 = "Q" & DatePart("q", tgt_append_Date) & "-" & Right(DatePart("yyyy", tgt_append_Date), 2)
    'if the new appending point is at a new Quarter sheet, then close the current one and loop to the new sheet
    'otherwise, exit the loop
    If QtrYr2 <> QtrYr Then
        QAwb.Close savechanges:=False
        QtrYr_Col = MyWs.Rows(1).Find(QtrYr2).Column
        'get the QA sheet
        QAsheet = MyWs.Cells(siteRow, QtrYr_Col).Value
    Else
        Exit Do
    End If
DoEvents
Loop
    'close the last QA sheet
    QAwb.Close savechanges:=False
    'get the finalized end date (will be used to determine rainfall end time)
    tgt_end_Row = Flws.Cells(Flws.Rows.count, Fl_Month_Col).End(xlUp).Row
    tgt_end_Date = DateSerial(Flws.Cells(tgt_end_Row, Fl_Year_Col).Value, Flws.Cells(tgt_end_Row, Fl_Month_Col).Value, Flws.Cells(tgt_end_Row, Fl_Day_Col).Value) + TimeSerial(Flws.Cells(tgt_end_Row, Fl_Hour_Col).Value, Flws.Cells(tgt_end_Row, Fl_Minute_Col).Value, 0)
    
'PART III : append rainfall
'open FP file on local drive
Dim RFwb As Workbook
If IsWorkBookOpen(target1) = False Then
    Set RFwb = Workbooks.Open(fileName:=target1)
Else
    Set RFwb = Workbooks(Dir(target1))
End If

'adjust windows for visual check
With RFwb.Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth * 0.5
    .Height = Application.UsableHeight * 0.6
    .Top = 0
    .Left = 0
    .ScrollRow = 1
    .ScrollColumn = 1
End With

Dim RFws As Worksheet
Set RFws = RFwb.Sheets(1)

With RFws.Rows(1)
    RF_Month_Col = .Find("Month", MatchCase:=False).Column
    RF_Day_Col = .Find("Day", MatchCase:=False).Column
    RF_Year_Col = .Find("Year", MatchCase:=False).Column
    RF_Hour_Col = .Find("Hour", MatchCase:=False).Column
    RF_Minute_Col = .Find("Minute", MatchCase:=False).Column
    RF_rainfall_Col = .Find("Rainfall", MatchCase:=False).Column
End With

RF_end_Row = RFws.Cells(RFws.Rows.count, RF_Month_Col).End(xlUp).Row
RF_end_Date = DateSerial(RFws.Cells(RF_end_Row, RF_Year_Col).Value, RFws.Cells(RF_end_Row, RF_Month_Col).Value, RFws.Cells(RF_end_Row, RF_Day_Col).Value)

If RF_end_Date < tgt_end_Date Then
    'Set source
    MyConn = "C:\Rainfall\RadarRain_TempFlowMon_Sheds_(since_2004).accdb"
     'Create query
    sSQL = "SELECT Month, Day, Year, Hour, Minute, [Rainfall(in)] FROM [RadarRain_TempFlowMon_Sheds_(since 2004)] WHERE (DTime >= #" & RF_end_Date & "# And Dtime <= #" & tgt_end_Date & "# and RainGauge='RG_" & siteName & "');"
    
     'Create RecordSet
    Set Cn = New ADODB.Connection
    With Cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"  'ACE is a newer and better oleDB driver than JET
       '.Provider = "Microsoft.Jet.OLEDB.4.0"
        .CursorLocation = adUseClient
        .Open MyConn
        Set rs = .Execute(sSQL)
    End With
    
    'Write RecordSet to results area
    RFws.Cells(RF_end_Row + 1, 1).CopyFromRecordset rs
    
    'release the object
    rs.Close
    Cn.Close
    Set Cn = Nothing
End If

'part IV: modify SSOAP database

'Hao Zhang @ 2015.4.15
'copy SSOAP template to destination, change file name
If fso.FileExists(target3) = False Then
    fso.CopyFile source3, target3
End If

Dim dbs As DAO.Database
Dim tdf As DAO.TableDef
Set dbs = OpenDatabase(target3)

'clear contents of all tables with a few exceptions
For Each tdf In dbs.TableDefs
    If Not (tdf.Name Like "*Units" Or tdf.Name Like "Holidays" Or tdf.Name Like "Metadata" Or tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
        dbs.Execute "delete * from " & tdf.Name
    End If
Next

start_time = DateSerial(Flws.Cells(2, Fl_Year_Col).Value, Flws.Cells(2, Fl_Month_Col).Value, Flws.Cells(2, Fl_Day_Col).Value) + TimeSerial(Flws.Cells(2, Fl_Hour_Col).Value, Flws.Cells(2, Fl_Minute_Col).Value, 0)
end_time = tgt_end_Date
DrainArea = MyWs.Cells(siteRow, DrainArea_Col).Value
'update contents of 5 tables by execute following queries
dbs.Execute "INSERT INTO Raingauges (RaingaugeID,RaingaugeName,RaingaugeLocationX,RaingaugeLocationY,RainUnitID,TimeStep,StartDateTime,EndDateTime) VALUES (1,'" & siteName & "', 0, 0, 1, 15,#" & start_time & "#,#" & end_time & "#);"
dbs.Execute "INSERT INTO RainConverters (RainConverterID,RainConverterName,RainUnitID,Format,LinesToSkip,MonthColumn,MonthWidth,DayColumn,DayWidth,YearColumn,YearWidth,HourColumn,HourWidth,MinuteColumn,MinuteWidth,RainColumn,RainWidth,CodeColumn,CodeWidth,MilitaryTime,AMPMColumn) VALUES (1,'" & siteName & "', 1, 'CSV',1,1,2,2,2,3,4,4,2,5,2,6,8,0,0,True,7);"
dbs.Execute "INSERT INTO Meters (MeterID, MeterName, StartDateTime, EndDateTime,Timestep, FlowUnitID, Area) VALUES (1,'" & siteName & "',#" & start_time & "#,#" & end_time & "#, 15, 1," & DrainArea & ");"
dbs.Execute "INSERT INTO FlowConverters (FlowConverterID,FlowConverterName,FlowUnitID,Format,LinesToSkip,MonthColumn,MonthWidth,DayColumn,DayWidth,YearColumn,YearWidth,HourColumn,HourWidth,MinuteColumn,MinuteWidth,FlowColumn,FlowWidth,CodeColumn,CodeWidth,MilitaryTime) VALUES (1,'" & siteName & "', 1, 'CSV',1,1,2,2,2,3,4,4,2,5,2,6,8,0,0,True);"
dbs.Execute "INSERT INTO Analyses (AnalysisID,AnalysisName,MeterID,RainGaugeID,BaseFlowRate,MaxDepressionStorage,RateOfReduction,InitialValue,R1,R2,R3,t1,T2,T3,K1,K2,K3,RunningAverageDuration,SundayDWFAdj,MondayDWFAdj,TuesdayDWFAdj,WednesdayDWFAdj,ThursdayDWFAdj,FridayDWFAdj,SaturdayDWFAdj,MaxDepressionStorage2,RateOfReduction2,InitialValue2,MaxDepressionStorage3,RateOfReduction3,InitialValue3) VALUES (1,'" & siteName & "_" & Format(Now(), "YYMMDD") & "_" & ini & "',1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0);"

Set dbs = Nothing
Application.DisplayAlerts = False
Application.ScreenUpdating = True

End Sub
Sub SSOAPcomplete(siteName As String)
'Hao Zhang @ 2015.3.13
'1. create Rainfall and FLOWINPUT files in local drive
'2. import data to each file
'get the titles, this is the fancy way, which can be used if varied columns are involved


'PART 0 : Initialization
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'get the destination folder to save generated files
Dim localpath As String
localpath = SSOAPpathTB.Value

'get the row number in CurSitesLB
siteRow = MyWs.Columns(siteName_Col).Find(siteName).Row
   
'PART I :file handling
Dim ini As String
If IsNull(SSOAPiniTB.Text) = False Then
    ini = SSOAPiniTB.Text
Else
    ini = InputBox("please enter your initial.")
End If

'set the full path of target files
Dim target1 As String
Dim target2 As String
Dim target3 As String
Dim source3 As String
tgtfldr = localpath & "\" & siteName & "\" & "SSOAP_" & Format(Now(), "yymmdd") & "_" & ini & "\"
target1 = tgtfldr & "Input\" & siteName & "_Rainfall_Radar2Sheds_" & Format(Now(), "yymmdd") & "_" & ini & ".csv"
target2 = tgtfldr & "Input\" & siteName & "_SSOAP_FLOWINPUT_" & Format(Now(), "yymmdd") & "_" & ini & ".csv"
source3 = "M:\Data\Temporary Monitors\Flow Monitoring\Supplementary Documents\QAQC_SSOAP_R-value_Templates\SSOAP Template\SSOAP Template (with holidays 2007-2016).sdb"
target3 = tgtfldr & siteName & "_SSOAP_" & Format(Now(), "yymmdd") & "_" & ini & ".sdb"

'if site subfolder is not exist, create the folder first
If fso.FolderExists(localpath & "\" & siteName & "\") = False Then
    fso.CreateFolder localpath & "\" & siteName & "\"
End If

'if site subfolder exists but the folder for current SSOAP run is not exist
If fso.FolderExists(tgtfldr) = False Then
    fso.CreateFolder tgtfldr
Else
    'clear all contents in the tgtfldr
    fso.DeleteFile tgtfldr & "*"
End If

If fso.FolderExists(tgtfldr & "input") = False Then
    fso.CreateFolder tgtfldr & "input"
Else
    'clear all contents in the tgtfldr
    fso.DeleteFile tgtfldr & "input\*"
End If

'PART II: import Flow data: recursively from each QA sheet

'set sheet numbers in new workbook to 1 (global variable)
Application.SheetsInNewWorkbook = 1
Set Flwb = Workbooks.Add
With Flwb
    'add heading
    .Sheets(1).range("A1:F1") = Split("Month, Day, Year, Hour, Minute, Flow", ",")
    .SaveAs fileName:=target2, FileFormat:=xlCSV, CreateBackup:=False
End With

With Flwb.Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth * 0.5
    .Height = Application.UsableHeight * 0.6
    .Top = 0
    .Left = Application.UsableWidth * 0.5
    .ScrollColumn = 1
    .ScrollRow = 1
    ''scroll to the end
    '.ScrollRow = FDwb.Sheets(1).Cells(1, 1).End(xlDown).Row - 10
End With

Dim Flws As Worksheet
Set Flws = Flwb.Sheets(1)
'get table defination
With Flws.Rows(1)
    Fl_Month_Col = .Find("Month", MatchCase:=False).Column
    Fl_Day_Col = .Find("Day", MatchCase:=False).Column
    Fl_Year_Col = .Find("Year", MatchCase:=False).Column
    Fl_Hour_Col = .Find("Hour", MatchCase:=False).Column
    Fl_Minute_Col = .Find("Minute", MatchCase:=False).Column
    Fl_flow_Col = .Find("Flow", MatchCase:=False).Column
End With

'get the first QA sheet
For QtrYr_Col = startQA_Col To endQA_Col
    If MyWs.Cells(siteRow, QtrYr_Col).Value <> "" Then
        QAsheet = MyWs.Cells(siteRow, QtrYr_Col).Value
        Exit For
    End If
Next

tgt_start_Row = 2
intvl = 15
'loop through each QA sheet, appending data to current QAQC file
Do While QAsheet <> ""
    'open QA_sheet
    If IsWorkBookOpen(QAsheet) = False Then
        Set QAwb = Workbooks.Open(fileName:=QAsheet, UpdateLinks:=False, ReadOnly:=True)
    Else
        Set QAwb = Workbooks(Dir(QAsheet))
    End If
    

    Set fd = QAwb.Sheets("Flow data")
    'get table defination
    With fd.range("10:20")
        src_dtime_Col = .Find("DateTime").Column
        src_Lvl_Col = .Find("Level 1").Column
        src_Corr_Flw_Col = .Find("Corrected Flow").Column
        'force the format of datetime column
        fd.Columns(src_dtime_Col).NumberFormat = "m/d/yyyy h:mm:ss"
        'find the first row of data, not date
        'the control structure solves the conditions that the first QAsheet don't start from the first day of the quarter
        If fd.Cells(.Find("Level 1").Row + 2, src_Lvl_Col).Value <> "" Then
            src_start_Row = .Find("Level 1").Row + 2
        Else
            src_start_Row = fd.Cells(.Find("Level 1").Row + 2, src_Lvl_Col).End(xlDown).Row
        End If
        src_start_Date = fd.Cells(src_start_Row, src_dtime_Col).Value
        src_end_Row = fd.Cells(fd.Rows.count, src_Lvl_Col).End(xlUp).Row
    End With
     
    tgt_end_Row = tgt_start_Row + src_end_Row - src_start_Row
     
    'put the date in an empty column (will be deleted later)
    Fl_dTime_Col = 7
    
    'copy new date
    If src_end_Row > src_start_Row Then
        With Flws
            'copy dTime to an empty column in tgt
            fd.range(.Cells(src_start_Row, src_dtime_Col).Address, .Cells(src_end_Row, src_dtime_Col).Address).Copy
            .range(.Cells(tgt_start_Row, Fl_dTime_Col).Address, .Cells(tgt_end_Row, Fl_dTime_Col).Address).PasteSpecial xlPasteValuesAndNumberFormats
            'stripping dTime into components
            .range(.Cells(tgt_start_Row, Fl_Month_Col).Address, .Cells(tgt_end_Row, Fl_Month_Col).Address).Formula = "=month(" & .Cells(tgt_start_Row, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            .range(.Cells(tgt_start_Row, Fl_Day_Col).Address, .Cells(tgt_end_Row, Fl_Day_Col).Address).Formula = "=day(" & .Cells(tgt_start_Row, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            .range(.Cells(tgt_start_Row, Fl_Year_Col).Address, .Cells(tgt_end_Row, Fl_Year_Col).Address).Formula = "=year(" & .Cells(tgt_start_Row, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            .range(.Cells(tgt_start_Row, Fl_Hour_Col).Address, .Cells(tgt_end_Row, Fl_Hour_Col).Address).Formula = "=hour(" & .Cells(tgt_start_Row, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            .range(.Cells(tgt_start_Row, Fl_Minute_Col).Address, .Cells(tgt_end_Row, Fl_Minute_Col).Address).Formula = "=minute(" & .Cells(tgt_start_Row, Fl_dTime_Col).Address(rowabsolute:=False, columnabsolute:=False) & ")"
            'remove formula by pasting the value
            tmpRng = .range(.Cells(tgt_start_Row, Fl_Month_Col).Address, .Cells(tgt_end_Row, Fl_Minute_Col).Address)
            .range(.Cells(tgt_start_Row, Fl_Month_Col).Address, .Cells(tgt_end_Row, Fl_Minute_Col).Address).Value = tmpRng
            'remove dTime column
            .range(.Cells(tgt_start_Row, Fl_dTime_Col).Address, .Cells(tgt_end_Row, Fl_dTime_Col).Address).Clear
            'paste flow data
            .range(.Cells(tgt_start_Row, Fl_flow_Col).Address, .Cells(tgt_end_Row, Fl_flow_Col).Address).Value _
             = fd.range(.Cells(src_start_Row, src_Corr_Flw_Col).Address, .Cells(src_end_Row, src_Corr_Flw_Col).Address).Value
        End With
    Else
        'break condition: when no more data could be added, exit the procedure
        Exit Do
    End If
    
QAwb.Close savechanges:=False

tgt_start_Row = tgt_end_Row + 1
QtrYr_Col = QtrYr_Col + 1
QAsheet = MyWs.Cells(siteRow, QtrYr_Col).Value
DoEvents
Loop
    
'PART III : append rainfall

'open FP file on local drive
Dim RFwb As Workbook
Set RFwb = Workbooks.Add
With RFwb
    'add heading
    .Sheets(1).range("A1:F1") = Split("Month, Day, Year, Hour, Minute, Rainfall", ",")
    .SaveAs fileName:=target1, FileFormat:=xlCSV, CreateBackup:=False
End With

'adjust windows for visual check
With RFwb.Windows(1)
    .WindowState = xlNormal
    .Width = Application.UsableWidth * 0.5
    .Height = Application.UsableHeight * 0.6
    .Top = 0
    .Left = 0
    .ScrollRow = 1
    .ScrollColumn = 1
End With

Dim RFws As Worksheet
Set RFws = RFwb.Sheets(1)

With RFws.Rows(1)
    RF_Month_Col = .Find("Month", MatchCase:=False).Column
    RF_Day_Col = .Find("Day", MatchCase:=False).Column
    RF_Year_Col = .Find("Year", MatchCase:=False).Column
    RF_Hour_Col = .Find("Hour", MatchCase:=False).Column
    RF_Minute_Col = .Find("Minute", MatchCase:=False).Column
    RF_rainfall_Col = .Find("Rainfall", MatchCase:=False).Column
End With

With Flws
    RF_start_Date = DateSerial(.Cells(2, Fl_Year_Col).Value, .Cells(2, Fl_Month_Col).Value, .Cells(2, Fl_Day_Col).Value)
    'get the finalized end date (will be used to determine rainfall end time)
    tgt_end_Row = .Cells(.Rows.count, Fl_Month_Col).End(xlUp).Row
    RF_end_Date = DateSerial(.Cells(tgt_end_Row, Fl_Year_Col).Value, .Cells(tgt_end_Row, Fl_Month_Col).Value, .Cells(tgt_end_Row, Fl_Day_Col).Value + 1)
End With

If RF_start_Date < RF_end_Date Then
    'Set source
    MyConn = "C:\Rainfall\RadarRain_TempFlowMon_Sheds_(since_2004).accdb"
     'Create query
    sSQL = "SELECT Month, Day, Year, Hour, Minute, [Rainfall(in)] FROM [RadarRain_TempFlowMon_Sheds_(since 2004)] WHERE (DTime >= #" & RF_start_Date & "# And Dtime <= #" & RF_end_Date & "# and RainGauge='RG_" & siteName & "');"
     'Create RecordSet
    Set Cn = New ADODB.Connection
    With Cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"  'ACE is a newer and better oleDB driver than JET
       '.Provider = "Microsoft.Jet.OLEDB.4.0"
        .CursorLocation = adUseClient
        .Open MyConn
        Set rs = .Execute(sSQL)
    End With
    
    'Write RecordSet to results area
    RFws.Cells(2, 1).CopyFromRecordset rs
    
    'release the object
    rs.Close
    Cn.Close
    Set Cn = Nothing
End If
    
'part IV: modify SSOAP database

'Hao Zhang @ 2015.4.15
'copy SSOAP template to destination, change file name
If fso.FileExists(target3) = False Then
    fso.CopyFile source3, target3
End If

DrainArea = MyWs.Cells(siteRow, DrainArea_Col).Value

Dim dbs As DAO.Database
Dim tdf As DAO.TableDef
Set dbs = OpenDatabase(target3)

'clear contents of all tables with a few exceptions
For Each tdf In dbs.TableDefs
    If Not (tdf.Name Like "*Units" Or tdf.Name Like "Holidays" Or tdf.Name Like "Metadata" Or tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
        dbs.Execute "delete * from " & tdf.Name
    End If
Next

'update contents of 5 tables by execute following queries
dbs.Execute "INSERT INTO Raingauges (RaingaugeID,RaingaugeName,RaingaugeLocationX,RaingaugeLocationY,RainUnitID,TimeStep,StartDateTime,EndDateTime) VALUES (1,'" & siteName & "', 0, 0, 1, 15,#" & rg_start_time & "#,#" & rg_end_time & "#);"
dbs.Execute "INSERT INTO RainConverters (RainConverterID,RainConverterName,RainUnitID,Format,LinesToSkip,MonthColumn,MonthWidth,DayColumn,DayWidth,YearColumn,YearWidth,HourColumn,HourWidth,MinuteColumn,MinuteWidth,RainColumn,RainWidth,CodeColumn,CodeWidth,MilitaryTime,AMPMColumn) VALUES (1,'" & siteName & "', 1, 'CSV',1,1,2,2,2,3,4,4,2,5,2,6,8,0,0,True,7);"
dbs.Execute "INSERT INTO Meters (MeterID, MeterName, StartDateTime, EndDateTime,Timestep, FlowUnitID, Area) VALUES (1,'" & siteName & "',#" & rg_start_time & "#,#" & rg_end_time & "#, 15, 1," & DrainArea & ");"
dbs.Execute "INSERT INTO FlowConverters (FlowConverterID,FlowConverterName,FlowUnitID,Format,LinesToSkip,MonthColumn,MonthWidth,DayColumn,DayWidth,YearColumn,YearWidth,HourColumn,HourWidth,MinuteColumn,MinuteWidth,FlowColumn,FlowWidth,CodeColumn,CodeWidth,MilitaryTime) VALUES (1,'" & siteName & "', 1, 'CSV',1,1,2,2,2,3,4,4,2,5,2,6,8,0,0,True);"
dbs.Execute "INSERT INTO Analyses (AnalysisID,AnalysisName,MeterID,RainGaugeID,BaseFlowRate,MaxDepressionStorage,RateOfReduction,InitialValue,R1,R2,R3,t1,T2,T3,K1,K2,K3,RunningAverageDuration,SundayDWFAdj,MondayDWFAdj,TuesdayDWFAdj,WednesdayDWFAdj,ThursdayDWFAdj,FridayDWFAdj,SaturdayDWFAdj,MaxDepressionStorage2,RateOfReduction2,InitialValue2,MaxDepressionStorage3,RateOfReduction3,InitialValue3) VALUES (1,'" & siteName & "_" & Format(Now(), "YYMMDD") & "_" & ini & "',1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0);"

Set dbs = Nothing

End Sub


Private Sub SSOAPrevertBtn_Click()
'Hao Zhang @ 2015.3.13
'discard created files

For Each wb In Application.Workbooks
    If wb.Name Like (siteName & "_SSOAP_FLOWINPUT*") Or wb.Name Like (siteName & "_Rainfall*") Then
        localpath = wb.path
        wbPath = wb.path & "\" & wb.Name
        wb.Close savechanges:=False
        Kill wbPath
    End If
Next

'ask if delete entire folder
'this if is to avoid error that a user may press the button with no worksheet opened
If fso.FolderExists(localpath) = True Then
    If MsgBox("Delete entire folder as well?", vbYesNo, "Warning") = vbYes Then
        'may have file system locking issue
        On Error Resume Next
        fso.DeleteFolder (fso.GetParentFolderName(localpath))
    End If
End If

End Sub

Private Sub SSOAPsaveBtn_Click()
'Hao Zhang @ 2015.3.13

'save created files
For Each wb In Application.Workbooks
    If wb.Name Like (siteName & "_SSOAP_FLOWINPUT*") Or wb.Name Like (siteName & "_Rainfall*") Then
        wb.Close savechanges:=True
    End If
Next

End Sub
Private Sub sdbBrowseBtn_Click()
'Hao Zhang @ 2015.4.23
'allow user to select the sdb file

'use this variable to save user's selection
tempPath = GetFile(sPath)
'if no path was specified, then retain the original value
If tempPath <> "" Then
    If fso.GetExtensionName(tempPath) = "sdb" Then
        sdbPathTB.Value = tempPath
    Else
        MsgBox "The file selected is not a .sdb file."
    End If
End If

End Sub

Private Sub sdbBtn_Click()

'Hao Zhang @ 4.23
'modifies existing SSOAP databases with updated info

Dim dbs As DAO.Database
Dim tdf As DAO.TableDef
If fso.FileExists(sdbPathTB.Text) = True Then
    NewName = fso.GetParentFolderName(sdbPathTB.Text) & "\" & SiteNameTB.Text & "_SSOAP_" & Format(Now(), "YYMMDD") & "_" & SSOAPiniTB.Text & ".sdb"
    'rename the file
    Name sdbPathTB.Text As NewName
    Set dbs = OpenDatabase(NewName)
Else
    MsgBox "Cannot open the file selected. Please check the path and try again."
End If
siteRow = MyWs.Columns(siteName_Col).Find(SiteNameTB.Text).Row
DrainArea = MyWs.Cells(siteRow, DrainArea_Col).Value
'clear contents of all tables with a few exceptions
For Each tdf In dbs.TableDefs
    If Not (tdf.Name Like "*Units" Or tdf.Name Like "Holidays" Or tdf.Name Like "Metadata" Or tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
        dbs.Execute "delete * from " & tdf.Name
    End If
Next

'update contents of 5 tables by execute following queries
dbs.Execute "INSERT INTO Raingauges (RaingaugeID,RaingaugeName,RaingaugeLocationX,RaingaugeLocationY,RainUnitID,TimeStep,StartDateTime,EndDateTime) VALUES (1,'" & SiteNameTB.Text & "', 0, 0, 1, 15,#" & sdbStartDP & "#,#" & sdbEndDP & "#);"
dbs.Execute "INSERT INTO RainConverters (RainConverterID,RainConverterName,RainUnitID,Format,LinesToSkip,MonthColumn,MonthWidth,DayColumn,DayWidth,YearColumn,YearWidth,HourColumn,HourWidth,MinuteColumn,MinuteWidth,RainColumn,RainWidth,CodeColumn,CodeWidth,MilitaryTime,AMPMColumn) VALUES (1,'" & SiteNameTB.Text & "', 1, 'CSV',1,1,2,2,2,3,4,4,2,5,2,6,8,0,0,True,7);"
dbs.Execute "INSERT INTO Meters (MeterID, MeterName, StartDateTime, EndDateTime,Timestep, FlowUnitID, Area) VALUES (1,'" & SiteNameTB.Text & "',#" & sdbStartDP & "#,#" & sdbEndDP & "#, 15, 1," & DrainArea & ");"
dbs.Execute "INSERT INTO FlowConverters (FlowConverterID,FlowConverterName,FlowUnitID,Format,LinesToSkip,MonthColumn,MonthWidth,DayColumn,DayWidth,YearColumn,YearWidth,HourColumn,HourWidth,MinuteColumn,MinuteWidth,FlowColumn,FlowWidth,CodeColumn,CodeWidth,MilitaryTime) VALUES (1,'" & SiteNameTB.Text & "', 1, 'CSV',1,1,2,2,2,3,4,4,2,5,2,6,8,0,0,True);"
dbs.Execute "INSERT INTO Analyses (AnalysisID,AnalysisName,MeterID,RainGaugeID,BaseFlowRate,MaxDepressionStorage,RateOfReduction,InitialValue,R1,R2,R3,t1,T2,T3,K1,K2,K3,RunningAverageDuration,SundayDWFAdj,MondayDWFAdj,TuesdayDWFAdj,WednesdayDWFAdj,ThursdayDWFAdj,FridayDWFAdj,SaturdayDWFAdj,MaxDepressionStorage2,RateOfReduction2,InitialValue2,MaxDepressionStorage3,RateOfReduction3,InitialValue3) VALUES (1,'" & SiteNameTB.Text & "_" & Format(Now(), "YYMMDD") & "_" & ini & "',1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0);"

Set dbs = Nothing


End Sub

'*************************************************************************
'****************************functions************************************
'*************************************************************************
Function IsWorkBookOpen2(fileName As String)
'Hao Zhang @ 2012.2.4
'check if a workbook is already open
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
Function IsWorkBookOpen(wbPath As String)
'Hao Zhang @ 2015.3.18
'simplified the function, wbPath must be the full path
For Each wb In Workbooks
    '.FullName returns the full path of the workbook
    If wb.FullName = wbPath Then
        IsWorkBookOpen = True
        Exit Function
    Else
        IsWorkBookOpen = False
    End If
Next
End Function

Private Function TabExists(tabName As String, wbName As Workbook)
'test if a tab (sheet) exists in a workbook
Dim ws As Worksheet
TabExists = False
For Each ws In wbName.Worksheets
    If ws.Name = tabName Then
        TabExists = True
        Exit Function
    End If
Next
End Function


Private Function GetUniqueName_Recursive(pathStr As String)
''Hao Zhang @ 2015.1.31
'Get a unique file name, including full path
'recursive method (not working properly yet, the loop method is used instead, see below)
If fso.FileExists(pathStr) = True Then
    flpath = fso.GetParentFolderName(pathStr) & "\"
    flname = fso.GetBaseName(pathStr)
    flExt = "." & fso.GetExtensionName(pathStr)

    If i = 0 Then
        i = 1
    End If

    If fso.FileExists(flpath & flname & " (" & i & ")" & flExt) = False Then
        GetUniqueName = flpath & flname & " (" & i & ")" & flExt
        Exit Function
    Else
        i = i + 1
        GetUniqueName_Recursive (flpath & flname & " (" & i & ")" & flExt)
    End If
Else
    GetUniqueName = pathStr
End Function

Private Function GetUniqueName(pathStr As String)
'Hao Zhang @ 2015.1.31
'Get a unique file name, including full path
'(alternative) loop method
If fso.FileExists(pathStr) = True Then
    flpath = fso.GetParentFolderName(pathStr) & "\"
    flname = fso.GetBaseName(pathStr)
    flExt = "." & fso.GetExtensionName(pathStr)
    i = 1
    Do
        i = i + 1
        GetUniqueName = flpath & flname & " (" & i & ")" & flExt
    Loop Until fso.FileExists(flpath & flname & " (" & i & ")" & flExt) = False
Else
    GetUniqueName = pathStr
End If
End Function

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

''decompose site name in Chosen Sites
'If siteName Like "*_?min" Then
'    intvl = Mid(siteName, InStr(1, siteName, "_") + 1, InStr(1, siteName, "min") - InStr(1, siteName, "_") - 1)
'    siteName = Left(siteName, InStr(1, siteName, "_") - 1)
'Else
'    intvl = 15
'    siteName = siteName
'End If

'loop through rows for each site
'*******************Should there is any change in the root Folder, change accordingly in here*****************
'    rPath = rootPath & siteName
'*************************************************************************************************************

'Set rFldr = fso.GetFolder(rPath)
    'find the site folders and write it into cells(iRow, 4)
'For Each Fldr In rFldr.SubFolders
'    'search the folders in the same level, 1 at a time
'    If InStr(1, Fldr.Name, "QAQC", vbTextCompare) <> 0 Then
'        QAQCFldr = Fldr
'        Exit For
'    End If

    'search the sub folders of the current folder, 1 at a time
    'use recursive to get into the ground level, then return to the upper level
'    If Fldr.SubFolders.Count > 0 Then
'Set Fldr = fso.GetFolder(rPath)
'QAQCFldr = RecurFldr(Fldr)
'    End If
'Next Fldr
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


'this part has not been implemented yet
'ErrorHandler:
'Dim errWs As Worksheet
'Dim MyWb As Workbook
'Set MyWb = ThisWorkbook
'Set errWs = MyWb.Sheets("ErrorLog")
'newRow = errWs.Cells(errWs.Rows.Count, 1).End(xlUp).Row + 1
'errWs.Cells(newRow, 1).Value = Now()
'errWs.Cells(newRow, 2).Value = "error on" & siteName
'Resume Next

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
Function GetFldr(strPath As String) As String
'Hao Zhang @ 2015.1.29
'returns a file's full path based on user's selection
Dim Fldr As FileDialog
Dim sItem As String
Set Fldr = Application.FileDialog(msoFileDialogFolderPicker)
With Fldr
    .title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFldr = sItem
Set Fldr = Nothing
End Function

Function GetFile(strPath As String) As String
'Hao Zhang @ 2015.1.29
'returns a file's full path based on user's selection
Dim fl As FileDialog
Dim sItem As String
Set fl = Application.FileDialog(msoFileDialogFilePicker)
With fl
    .title = "Select a File"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFile = sItem
Set fl = Nothing
End Function
Function Col_Letter(lngCol) As String
'convert column number into letter
'http://stackoverflow.com/questions/12796973/vba-function-to-convert-column-number-to-letter
Dim vArr
vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_Letter = vArr(0)
End Function

Private Sub UserForm_Terminate()
'save input info in cacheTbl before exit
With MyWb.Sheets("cacheTbl")
    'QAQC:
    .Cells(2, 2).Value = QAtempTB.Text
    .Cells(3, 2).Value = QAtempQtrYrCB.Text
    .Cells(4, 2).Value = QAtempIntvlTB.Text
    .Cells(5, 2).Value = QAqtrYrCB.Text
    .Cells(6, 2).Value = RawDateTB.Text
    .Cells(7, 2).Value = QAiniTB.Text
    .Cells(8, 2).Value = StartTimeTB.Text
    .Cells(9, 2).Value = EndTimeTB.Text
    .Cells(10, 2).Value = RGCB.Text
    'BigPicture:
    .Cells(12, 2).Value = BigPpathTB.Text
    .Cells(13, 2).Value = BigPfullOB
    .Cells(14, 2).Value = BigPiniTB.Text
    .Cells(15, 2).Value = BigPpptPathTB.Text
    .Cells(16, 2).Value = BigPpptYrCB.Text
    .Cells(17, 2).Value = BigPpptMonCB.Text
    'SSOAP:
    .Cells(19, 2).Value = SSOAPpathTB.Text
    .Cells(20, 2).Value = SSOAPfullOB
    .Cells(21, 2).Value = SSOAPiniTB.Text
    .Cells(22, 2).Value = sdbPathTB.Text
    .Cells(23, 2).Value = sdbStartDP.Value
    .Cells(24, 2).Value = sdbEndDP.Value
    .Cells(25, 2).Value = SiteNameTB.Text
End With

End Sub
