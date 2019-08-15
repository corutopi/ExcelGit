Attribute VB_Name = "SS204_ExcelSheetProperty"
Option Explicit


'*******************************************************************************
'--comment_version_0.1.0------------------
'選択したセルに入力規則がせていされているかどうかを判定する.
'複数の範囲のセルが指定されている場合は...どうしよう.
'
'
'
'-----------------------------------------
'引数       :hikihiki       :引数の説明
'戻り値     :戻り値の内容
'-----------------------------------------
'--更新履歴-------------------------------
'20181126   :xxx            :[更新内容]
'*******************************************************************************
Public Function HasVaridation(target As Range) As Boolean
    Const CP_TARGET_ERR As Long = 1004
    Dim re As Boolean
    Dim beforeErr As Long
    
    re = False
    beforeErr = Err.Number
    Err.Clear
    On Error Resume Next
    If target.Validation.Type Then
    End If
    If Err.Number = 0 Then
        re = True
    ElseIf Err.Number = CP_TARGET_ERR Then
        '想定のエラーであれば入力規則なし
        re = False
    Else
        'その他のエラーの場合は処理に失敗と判断
        On Error GoTo 0
        Err.Raise Err.Number
    End If
    If beforeErr <> 0 Then
        '処理実施前にエラーが発生していた場合は同じエラーが発生した状態にする
        Err.Raise beforeErr
    End If
    
    HasVaridation = re
End Function
