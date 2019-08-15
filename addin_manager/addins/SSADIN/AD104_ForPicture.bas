Attribute VB_Name = "AD104_ForPicture"
Option Explicit

Private pr_height As Double
Private pr_width As Double

Private pr_trimTop As Double
Private pr_trimLeft As Double
Private pr_trimBottom As Double
Private pr_trimRight As Double


'*******************************************************************************
'-----------------------------------------
'選択中のオブジェクト(またはセル範囲)のサイズを変数に記録する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171230   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_画像サイズ記憶()
    On Error GoTo ERR_MANAGER
    
    pr_height = Selection.Height
    pr_width = Selection.Width
    
    Debug.Print "高さ：" & pr_height
    Debug.Print "幅　：" & pr_width
    
    Exit Sub
    
ERR_MANAGER:
    Call errProcess("サイズが正しく取得できませんでした。" _
            & vbNewLine & "セルかオブジェクトを選択した状態で実行してください。")
End Sub


'*******************************************************************************
'-----------------------------------------
'「SA_画像サイズ記憶」で記憶したサイズを選択中のオブジェクトに反映する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171230   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_画像サイズ反映()
    On Error GoTo ERR_MANAGER
    
    If pr_height = 0 Or pr_width = 0 Then
        MsgBox "画像サイズ記憶を最初に行ってください。", vbInformation
    Else
        Selection.Height = pr_height
        Selection.Width = pr_width
    End If
    
    Exit Sub
    
ERR_MANAGER:
    Call errProcess("サイズが正しく反映されませんでした" _
            & vbNewLine & "オブジェクトを選択した状態で実行してください。")
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のオブジェクトのトリミング範囲を変数に記録する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180205   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_画像トリミング範囲記憶()
    On Error GoTo ERR_MANAGER
    
    With Selection.PictureFormat
        pr_trimTop = .CropTop
        pr_trimLeft = .CropLeft
        pr_trimBottom = .CropBottom
        pr_trimRight = .CropRight
    End With
    
    Exit Sub
    
ERR_MANAGER:
    Call errProcess("トリミング範囲が正しく取得できませんでした。" _
            & vbNewLine & "オブジェクトを選択した状態で実行してください。")
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のオブジェクトに記録したトリミング範囲を反映する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180205   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_画像トリミング反映()
    On Error GoTo ERR_MANAGER
    
    With Selection.PictureFormat
        .CropTop = pr_trimTop
        .CropLeft = pr_trimLeft
        .CropBottom = pr_trimBottom
        .CropRight = pr_trimRight
    End With
    
    Exit Sub
    
ERR_MANAGER:
    Call errProcess("サイズが正しく反映されませんでした" _
            & vbNewLine & "オブジェクトを選択した状態で実行してください。")
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中の画像を記憶する。
'
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180122   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_画像契約()
    Const CL_AGREE_BEAST As String = "契約獣01"
    Dim ws As Worksheet
    
    Dim target As Shape
    Dim s As Shape
    
    On Error GoTo ERR_MANAGER
    
    If Selection.ShapeRange.Count > 1 Then GoTo ERR_MANAGER
    
    Set ws = ThisWorkbook.Worksheets(1)
    Set target = Selection.ShapeRange(1)
    
    For Each s In ws.Shapes
        If s.Name = CL_AGREE_BEAST Then
            s.Delete
            Exit For
        End If
    Next
    
    target.Copy
    ws.Paste
    
    Debug.Print Selection.Name
    
    ws.Shapes(ws.Shapes.Count).Name = CL_AGREE_BEAST
    
    Exit Sub
    
ERR_MANAGER:
    MsgBox "契約に失敗しました。" _
            & vbNewLine & "単一のオブジェクトを選択して実行してください。", vbExclamation
End Sub


'*******************************************************************************
'-----------------------------------------
'契約した画像を貼り付ける。
'
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180122   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_契約画像召喚()
    Const CL_AGREE_BEAST As String = "契約獣01"
    
    On Error GoTo ERR_MANAGER
    ThisWorkbook.Worksheets(1).Shapes(CL_AGREE_BEAST).Copy
    ActiveSheet.Paste
    
    Exit Sub
    
ERR_MANAGER:
    MsgBox "召喚に失敗しました。" _
            & vbNewLine & "先に画像契約を行ってください。", vbExclamation
End Sub


