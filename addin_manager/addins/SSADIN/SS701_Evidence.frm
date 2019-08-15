VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SS701_Evidence 
   Caption         =   "海老蔵君"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3990
   OleObjectBlob   =   "SS701_Evidence.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SS701_Evidence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClipboardMonitoring_Change()
    If ClipboardMonitoring.Value = True Then
        Call catchClipboard(ActiveSheet)
        
        BookName.Caption = ActiveWorkbook.Name
        SheetName.Caption = ActiveSheet.Name
        
        Label2.Visible = True
        BookName.Visible = True
        SheetName.Visible = True
    Else
        Call releaseClipboard
        
        Label2.Visible = False
        BookName.Visible = False
        SheetName.Visible = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If ClipboardMonitoring.Value = True Then
        Call releaseClipboard
        ClipboardMonitoring.Value = False
    End If
End Sub

