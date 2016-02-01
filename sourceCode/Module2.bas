Attribute VB_Name = "Module2"
Sub Macro1()
'
' Macro1 Macro
' (recorded) change formula and color for 3 cells in S50-011535

'
    range("M28").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    range("M37").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    range("M41").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    range("O28").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[3]="""","""",(VLOOKUP(RC[-1],'Area vs. depth table'!C[-14]:C[-12],3,TRUE)-VLOOKUP(RC[3],'Area vs. depth table'!C[-14]:C[-12],3,TRUE)))"
    range("O28").Select
    range("O28").AddComment
    range("O28").Comment.visible = False
    range("O28").Comment.Text Text:="Hao Zhang:" & Chr(10) & "Switched to DU"
    range("O28").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    Selection.Copy
    range("O37").Select
    ActiveSheet.Paste
    range("O41").Select
    ActiveSheet.Paste
    range("O43").Select
End Sub


Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "m\n14"
'
' Macro2 Macro
'' (recorded) change value and color


    ActiveWorkbook.Sheets("Site Info").Activate

    range("B2").Value = Format(Now(), "m/d/yyyy")
    range("C2").Value = "HZ"
    range("B9").Value = "DU"
    range("B10").Value = "UU"
    range("C14").Value = "Acres"
    range("C16").Value = "PipeFlowAreas_20150225_HZ.xlsm"
    range("C17").Value = "M:\Data\Temporary Monitors\Flow Monitoring\Supplementary Documents"
    range("B20").Value = #2/4/2015#
    
    range("l29:l41").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Selection.Font.Color = xlblack
    
    range("l56").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Selection.Font.Color = xlblack
    
    range("m29").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[3]="""","""",(VLOOKUP(RC[-3],'Area vs. depth table'!C[-14]:C[-12],3,TRUE)-VLOOKUP(RC[3],'Area vs. depth table'!C[-14]:C[-12],3,TRUE)))"
    range("m29").Select
    range("m29").AddComment
    range("m29").Comment.visible = False
    range("m29").Comment.Text Text:="Hao Zhang:" & Chr(10) & "Switched to UU"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    Selection.Copy
    range("m30:m41").Select
    ActiveSheet.Paste
    range("m56").Select
    ActiveSheet.Paste
'    Range("T38").Select
'    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .Color = 255
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With
'    Selection.ClearContents
End Sub

Sub test()
Date = Now()
Debug.Print Format(Now(), "m\m\i\n s\s\e\c")

End Sub
