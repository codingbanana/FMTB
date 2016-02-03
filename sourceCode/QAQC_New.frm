VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QAQC_New 
   Caption         =   "QAQC_New"
   ClientHeight    =   6855
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6204
   OleObjectBlob   =   "QAQC_New.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QAQC_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()
Call extendFPrange
End Sub

Private Sub QuarterCbox_Change()

End Sub

Private Sub TrimTailBtn_Click()
Call trim_tail
End Sub

Private Sub UserForm_Initialize()
With QuarterCbox
        .AddItem "Q1 (Jan-Mar)"
        .AddItem "Q2 (Apr-Jun)"
        .AddItem "Q3 (Jul-Sept)"
        .AddItem "Q4 (Oct-Dec)"
End With

With yearCbox
        .AddItem "2014"
        .AddItem "2015"
        .AddItem "2016"
        .AddItem "2017"
End With

startTimeTextBox = #12/2/2014#
EndTimeTextBox = #1/2/2015#
RGTextBox = 1
End Sub

Private Sub CommandButton1_Click()
Call QC_Year_shift
End Sub

Private Sub CommandButton2_Click()
Call Month_TS_chart
End Sub

Private Sub CommandButton3_Click()
Call temp_chart
End Sub


Private Sub monthCbox_Change()
Select Case QAQC_form.QuarterCbox.Value
Case Is = "Q1 (Jan-Mar)"
QAQC_form.startTimeTextBox.Value = DateSerial(QAQC_form.yearCbox.Value, 1, 1)
QAQC_form.EndTimeTextBox.Value = DateSerial(QAQC_form.yearCbox.Value, 4, 2)
Case Is = "Q2 (Apr-Jun)"
QAQC_form.startTimeTextBox.Value = DateSerial(QAQC_form.yearCbox.Value, 4, 1)
QAQC_form.EndTimeTextBox.Value = DateSerial(QAQC_form.yearCbox.Value, 7, 2)
Case Is = "Q3 (Jul-Sept)"
QAQC_form.startTimeTextBox.Value = DateSerial(QAQC_form.yearCbox.Value, 7, 1)
QAQC_form.EndTimeTextBox.Value = DateSerial(QAQC_form.yearCbox.Value, 10, 2)
Case Is = "Q4 (Oct-Dec)"
QAQC_form.startTimeTextBox.Value = DateSerial(QAQC_form.yearCbox.Value, 10, 1)
QAQC_form.EndTimeTextBox.Value = DateSerial(QAQC_form.yearCbox.Value + 1, 1, 2)
End Select
End Sub

Private Sub rainlinkButton_Click()
Call RainlinkNew
End Sub



Private Sub yearCbox_Change()
Call monthCbox_Change
End Sub

