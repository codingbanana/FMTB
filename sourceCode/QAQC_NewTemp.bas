Attribute VB_Name = "QAQC_NewTemp"
Sub RainlinkNew()
'Hao Zhang @2015.1.14
'updated cell reference to adapt new template
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
rainCol = fd.range("A12:AZ12").Find("Rain Fall Data").Column
'convert the column number back to letter
vArr = Split(Cells(1, rainCol).Address(True, False), "$")
'link Flow data!rainfall data (column AC) to Rainfall!rainfall(column B)

fd.Activate
fd.Cells(14, rainCol).Select
If Selection.Formula <> "='Rainfall Data'!B2" Then
    Selection.Value = "='Rainfall Data'!B2"
    Selection.AutoFill Destination:=range(vArr(0) & "14", vArr(0) & EOR)
End If

End Sub


