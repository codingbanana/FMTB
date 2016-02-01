Attribute VB_Name = "SSOAP"

Sub rain_Query()
'
' generate a rainfall table by querying the PWD2010.mdb
' user need to specify RG#, start time, end time
'
    Dim RG As Integer
    Dim startTime As Date, endTime As Date

    Dim Cn As ADODB.Connection, rs As ADODB.Recordset
    Dim MyConn, sSQL As String
    
    Dim iCol
    Dim ws As Worksheet
    
   
    Dim MyField, Location As range  'This is the tricky part, MyField is actually dimed as Variant
    
    Application.ScreenUpdating = False
    
    Dim t0, t1
    'set up a timer, can also use now()
    t0 = Timer
    
    ' define default worksheet
   Set ws = Worksheets("Rain")
         Windows("Rainfall_Flow_Dtime_Convert.xlsx").Activate
         ws.Select
         RG = range("K1").Value
         startTime = range("K2").Value
         endTime = range("K3").Value
      
      'Set destination
    Set Location = [A2]
     'Set source
    MyConn = "C:\Rainfall\PWDRAIN2010\PWDRAIN2010.mdb"
     'Create query
    sSQL = "SELECT Daytime, finalRG" & RG & " FROM [FinalAll(2014)] WHERE (((Daytime) >= #" & startTime & "# And (Daytime) <=#" & endTime & "#));"
    
     'Create RecordSet
    Set Cn = New ADODB.Connection
    With Cn
    
        .Provider = "Microsoft.ACE.OLEDB.12.0"  'ACE is a newer and better oleDB driver than JET
'       .Provider = "Microsoft.Jet.OLEDB.4.0"
        .CursorLocation = adUseClient
        .Open MyConn
        Set rs = .Execute(sSQL)
    End With
    'get the title
    For icols = 0 To rs.Fields.count - 1
    ws.Cells(1, icols + 1).Value = rs.Fields(icols).Name
    Next
    'set title font to be bold
    ws.range(ws.Cells(1, 1), ws.Cells(1, rs.Fields.count)).Font.Bold = True
    'Clear previous results
    range("A2:B2").Select
        range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
    'Write RecordSet to results area
        range("A2").CopyFromRecordset rs
    
    
   ''the copyFromRecordset method is MUCH MUCH faster than the loop method below
'      ReDim Result(1 To rs.RecordCount, 1 To rs.Fields.Count)
'    Rw = Location.Row
'    Col = Location.Column
'   c = Col
'   Do Until rs.EOF
'      For Each MyField In rs.Fields
'          Cells(Rw, c) = MyField
'           c = c + 1
'       Next MyField
'       rs.MoveNext
'       Rw = Rw + 1
'       c = Col
'       'set a max running time = 51s (0.0006)
'       t1 = Now()
'           If (t1 - t0) > 0.0006 Then
'           MsgBox "Macro has terminated prematurely due to unexpected long run time." & Chr(13) & "" & Chr(10) & "Sorry, please run the macro again", vbExclamation, "Error"
 '          Exit Sub
'           End If
 '  Loop
    rs.Close
    Cn.Close
    Set Location = Nothing
    Set Cn = Nothing
    
'clear the previous autofilled cells and refill to the last row
    range("C3:H3").Select
        range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
    range("C2:H2").AutoFill Destination:=range("C2:H" & range("B" & Rows.count).End(xlUp).Row)
    
    'clean up the extra rows on bottom
   ' If Cells(Rw, 3).Value <> "" Then
   ' Range("C" & Rw & ":H" & Rw).Select
   ' Range(Selection, Selection.End(xlDown)).Select
   ' Selection.ClearContents
   ' End If
    
    range("C2:H2").Select
     range(Selection, Selection.End(xlDown)).Select
   
   t1 = Round(Timer - t0, 0)
    'MsgBox "RG#" & RG & " data from" & startTime & " to" & endTime & " have been imported." & Chr(13) & "" & Chr(10) & "Total imported entries: " & Rw & ". Run Time is " & Minute(t1 - t0) & ":" & Second(t1 - t0) & Chr(13) & "" & Chr(10) & "Thanks for using this Macro, Have a good day!", vbInformation, "Success!"
    MsgBox "done in " & t1 & " seconds!"
End Sub
Sub rain_query_qk()
'Hao Zhang @ 2014/12/19
'the Access-Excel ETL is too time consuming sometimes
'a simplified proc that pulls data from AllRG tab is developed
Dim i As Integer
Dim startTime As Date
Dim endTime As Date
Dim RG As Integer
RG = Sheets("Rain").range("k1").Value
startTime = Sheets("Rain").range("K2").Value
endTime = Sheets("Rain").range("K3").Value
Date_Col = 1
rain_col = Sheets("AllRG").range("B1:AJ1").Find("finalRG" & RG).Column
'direct find datetime will result in runtime error 91, which means no value could be found
'trick is to convert datetime into double floating numbers using CDbl()
first_row = Sheets("AllRG").range("A2:A30000").Find(CDbl(startTime)).Row
last_row = Sheets("AllRG").range("A2:A30000").Find(CDbl(endTime)).Row
Length = last_row - first_row
Sheets("Rain").range("A2:A" & (2 + Length)).Value = Sheets("AllRG").range(Cells(first_row, Date_Col), Cells(last_row, Date_Col)).Value
Sheets("Rain").range("B2:B" & (2 + Length)).Value = Sheets("AllRG").range(Cells(first_row, rain_col), Cells(last_row, rain_col)).Value

End Sub
Sub Pull_Data_from_Excel_with_ADODB()
'query an Excel table using SQL way
'from stackoverflow.com
'not working yet, need more tuning
    Dim cnStr As String
    Dim rs As ADODB.Recordset
    Dim query As String

    Dim fileName As String
    fileName = "C:\Users\hao.zhang\Desktop\Rainfall_Flow_Dtime_Convert.xlsx"

    cnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
               "Data Source=" & fileName & ";" & _
               "Extended Properties=Excel 12.0"

    'query = "SELECT * FROM [Sheet]"
    query = "SELECT * FROM [AllRG] WHERE [Daytime] > #10/1/2014# and [daytime] < #12/1/2014#"

    Set rs = New ADODB.Recordset
    rs.Open query, cnStr, adOpenUnspecified, adLockUnspecified

    Cells.Clear
    range("A2").CopyFromRecordset rs

   ' Dim cell As Range, i As Long
   ' With Range("A1").CurrentRegion
   '     For i = 0 To rs.Fields.Count - 1
    '        .Cells(1, i + 1).Value = rs.Fields(i).Name
   ''     Next i
   '     .EntireColumn.AutoFit
   ' End With
End Sub

Sub flow_query()
Dim site As String

site = range("K1").Value

Workbooks("Rainfall_Flow_Dtime_Convert.xlsx").Activate
Worksheets("Flow").Select
range("B2:B30000").Clear
Workbooks(site & " (Q3-14).xlsm").Activate
Worksheets("Flow Data").Select
range("W2991:W8846").Copy
Workbooks("Rainfall_Flow_Dtime_Convert.xlsx").Activate
Worksheets("Flow").Select
range("B2").PasteSpecial xlPasteValues
Workbooks(site & " (Q4-14).xlsm").Activate
Worksheets("Flow Data").Select
range("W15:W5870").Copy
Workbooks("Rainfall_Flow_Dtime_Convert.xlsx").Activate
Worksheets("Flow").Select
range("B5858").Select
ActiveCell.PasteSpecial xlPasteValues
range("C2", "H11713").Select
End Sub

