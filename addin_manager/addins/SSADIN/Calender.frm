VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calender 
   Caption         =   "UserForm1"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5190
   OleObjectBlob   =   "Calender.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Calender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pr_LblTable(1 To 42) As MSForms.Label



Private Sub makeCalender()
    Dim y As Long
    Dim m As Long
    Dim startLabelNum As Long
    Dim tmpDate As Date
    Dim i As Long
    
    y = 2018
    m = 1
    
    startLabelNum = Weekday(DateSerial(y, m, 1))
    For i = 1 To 42
        If i < startLabelNum Or startLabelNum + Day(DateSerial(y, m + 1, 0)) - 1 < i Then
            pr_LblTable(i).Visible = False
        Else
            tmpDate = DateSerial(y, m, i - startLabelNum + 1)
            pr_LblTable(i).Visible = True
            pr_LblTable(i).Caption = Right(" " & Day(tmpDate), 2)
            pr_LblTable(i).ForeColor = getDayColor(tmpDate)
        End If
    Next
End Sub

Private Function getDayColor(d As Date)
    Dim re As Long
    'デフォ
    If Weekday(d, vbMonday) >= 6 Then
        re = vbRed
    Else
        re = vbBlack
    End If
    
    getDayColor = re
End Function

Private Sub UserForm_Initialize()
    Dim i As Long
    
    For i = 1 To 42
        Set pr_LblTable(i) = Me.Controls("Label" & i)
    Next
    
    Call makeCalender
End Sub
