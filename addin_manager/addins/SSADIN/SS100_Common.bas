Attribute VB_Name = "SS100_Common"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub startProcess(Optional ByVal Name As String = "")
    Application.ScreenUpdating = False
    Call logging("処理開始：" & Name)
End Sub

Public Sub endProcess(Optional ByVal Name As String = "")
    Call logging("処理終了：" & Name)
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub logging(ByVal msg As String)
    Application.StatusBar = msg
End Sub

Public Sub errProcess(Optional ByVal Description As String = "", Optional ByVal isEnd As Boolean = False)
    Dim msg As String
    
    If Description <> "" Then
        'エラー文言が個別に設定されている場合
        msg = Description
    ElseIf Err.Number <> 0 Then
        'システムエラーで呼び出された場合
        msg = "エラーが発生しました。" _
                & vbNewLine & "エラー番号：" & Err.Number _
                & vbNewLine & "エラー内容：" & Err.Description
    Else
        'エラー文言がなくシステムエラーでもない場合
        msg = "エラー判定が行われましたが、詳細が判別できませんでした。"
    End If
    
    MsgBox msg, vbExclamation, "共通エラー処理"
    
    If isEnd Then
        End     '強制終了
    End If
End Sub

Public Sub errTest()
    Debug.Print "start"
    
    Call errProcess
    
    Debug.Print "end"
End Sub
