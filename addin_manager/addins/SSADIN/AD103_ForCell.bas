Attribute VB_Name = "AD103_ForCell"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'選択中のセルに対し入力されている値を再入力する。
'
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180122   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_セル値再入力()
    Dim r As Range
    
    Dim tmp As String
    Dim i As Long
    
    Set r = Selection
    
    For i = 1 To r.Count
        tmp = r.Item(i).Formula
        r.Item(i).Formula = tmp
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のセルを表としてフィルターなどをつけたレイアウトに加工する。
'
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20181219   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_表レイアウト化()
    Const HEADER_COLOR = 16777185
    Dim target As Range
    
    '初期化
    Set target = Selection
    'レイアウトリセット
    target.Borders.LineStyle = xlNone
    target.Font.Bold = False
    target.Interior.Pattern = xlNone
    '罫線
    With target.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    'ヘッダー編集
    target.Rows(1).Interior.Color = HEADER_COLOR
    target.Rows(1).Font.Bold = True
    'フィルター設定
    If Not target.Parent.AutoFilter Is Nothing Then
        '既存のフィルターの解除
        target.AutoFilter
    End If
    target.AutoFilter
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のセルのセル結合を解除して結合前の値で解除後のセルすべてを埋める。
'
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20181219   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_DB風データ_結合充足()
    Dim target As Range
    Dim tmpMergeArea As Range
    Dim tmpValue As Variant
    Dim i As Long
    Dim j As Long
    
    Set target = Selection
    For i = 1 To target.Count
        If target.Item(i).MergeCells Then
            Set tmpMergeArea = target.Item(i).MergeArea
            tmpValue = target.Item(i).Value
            tmpMergeArea.UnMerge
            For j = 2 To tmpMergeArea.Count
                tmpMergeArea.Item(j).Value = tmpValue
            Next
        End If
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のセルの空白セルを1つ上のセル値で埋める
'
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20181219   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_DB風データ_空欄充足()
    Dim target As Range
    Dim r As Long
    Dim c As Long
    
    Set target = Selection
    For c = 1 To target.Columns.Count
        For r = 2 To target.Rows.Count
            If target.Item(r, c).Value = "" Then
                target.Item(r, c).Value = target.Item(r - 1, c).Value
            End If
        Next
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のセル範囲のすべての行を指定した行数ずつ開ける.
'セルは下方向に挿入移動する.
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20181219   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_指定行ずつ挿入()
    Dim target As Range
    Dim tmpSU As Boolean
    Dim tmpStr As String
    Dim interval As Integer
    Dim i As Long
    
    '初期化
    tmpSU = ChangeScreenUpdate(False)
    '挿入列数取得
    Do
        tmpStr = InputBox("挿入する行数を指定してください.")
        If tmpStr = "" Then
            Exit Sub
        ElseIf Not IsNumeric(tmpStr) Then
            MsgBox "整数を指定してください.", vbExclamation
        ElseIf Not CDbl(tmpStr) = Int(tmpStr) Then
            MsgBox "整数を指定してください.", vbExclamation
        Else
            interval = CInt(tmpStr)
            Exit Do
        End If
        tmpStr = ""
    Loop While tmpStr = ""
    'データ挿入
    Dim ws As Worksheet
    Dim insertRange As Range
    Dim tlRange As Range
    Dim brRange As Range
    
    Set ws = ActiveSheet
    Set target = Selection
    For i = 1 To target.Rows.Count - 1  '最終行は対象外
        Set tlRange = ws.Cells(target.Item(1).row - 1 + 2 + (interval + 1) * (i - 1), _
                               target.Item(1).column)
        Set brRange = ws.Cells(target.Item(1).row - 1 + 2 + (interval + 1) * (i - 1) + (interval - 1), _
                               target.Item(target.Count).column)
        Set insertRange = ws.Range(tlRange, brRange)
        insertRange.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    '終了処理
    ChangeScreenUpdate tmpSU
End Sub














