Attribute VB_Name = "QCpullinfo"
Dim fso As New FileSystemObject
Sub getDrainArea()

'Hao Zhang @ 2015.2.17
'get site drainage area from QA sheets
Set MyWs = ThisWorkbook.Worksheets("CurSitesTbl")
Set logws = Workbooks("QA Logbook.xlsm").Worksheets("DrainageArea")
i = 2
Do While MyWs.Cells(i, 22).Value <> ""
Set QAwb = Workbooks.Open(fileName:=MyWs.Cells(i, 22).Value, UpdateLinks:=False, ReadOnly:=True)
QAwb.Windows(1).visible = False
With QAwb.Worksheets("Site Info")
    logws.Cells(i, 1) = .range("1:30").Find("Site Name:").Offset(0, 1).Value
    logws.Cells(i, 2) = .range("1:30").Find("Drainage Area (acre):").Offset(0, 1).Value
    i = i + 1
End With
Application.DisplayAlerts = False
QAwb.Close
DoEvents
Loop
MsgBox "Done."
End Sub

Private Sub getPercentRecovery()
'Hao Zhang @ 2015.2.22
'fill yellow sheet with %recovery

'yellowSheet = "M:\Data\Temporary Monitors\Flow Monitoring\Supplementary Documents\CSL Meeting Notes\2015\February 2015\Flow Monitoring % Recovery_January -HZ.xlsx"
Set MyWs = ThisWorkbook.Worksheets("CurSitesTbl")

yellowSheet = "C:\Users\hao.zhang\Desktop\Flow Monitoring % Recovery_January -HZ.xlsx"

Set yswb = Workbooks.Open(yellowSheet)
    
With MyWs
Dim p As Integer
    For N = 2 To 62
        wbPath = fso.GetParentFolderName(.Cells(N, 23).Value)
        wbFile = fso.GetFileName(.Cells(N, 23).Value)
        wbsheet = "Flow Data"
        wbrng = range("I5").Address
        site = MyWs.Cells(N, 4).Value
        For p = 1 To 3
            wbref = "='" & wbPath & "\[" & wbFile & "]" & wbsheet & "'!" & wbrng
            yswb.Sheets(1).Columns(1).Find(site).Offset(0, 3 + p).Formula = wbref
            yswb.Sheets(1).Columns(1).Find(site).Offset(0, 3 + p).Value = yswb.Sheets(1).Columns(1).Find(site).Offset(0, 3 + p).Value
            wbrng = range(wbrng).Offset(0, 1).Address
            
        Next
    Next
End With
End Sub

