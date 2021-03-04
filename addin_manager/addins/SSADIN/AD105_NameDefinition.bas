Attribute VB_Name = "AD105_NameDefinition"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'選択中のブック内の名前定義を選択中のセルから下方向に書き出す。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20210305   :sueki          : add single quate before NameDefinition name
'                           : for Sheet names that inclued symbols.
'20180318   :sueki          : create
'*******************************************************************************
Public Sub SA_名前定義書き出し()
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '実行確認
    If MsgBox("選択中のセルから下方向に名前定義の情報を出力します" _
                & vbNewLine & "よろしいですか？", vbYesNo) = vbNo Then Exit Sub
    
    r = Selection.Item(1).row
    c = Selection.Item(1).column
    For Each n In ActiveWorkbook.names
        ActiveSheet.Cells(r, c + 0).Value = "'" & n.Name
        ActiveSheet.Cells(r, c + 1).Value = "'" & n.RefersTo
        ActiveSheet.Cells(r, c + 2).Value = "'" & n.Comment
        r = r + 1
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のシート内の名前定義を選択中のセルから下方向に書き出す。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20210305   :sueki          : add single quate before NameDefinition name
'                           : for Sheet names that inclued symbols.
'20180318   :sueki          : create
'*******************************************************************************
Public Sub SA_名前定義書き出し_シートオンリー()
    Dim s As Worksheet
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '実行確認
    If MsgBox("選択中のセルから下方向に、選択中のシートに参照範囲が定義されている名前定義の情報を出力します" _
                & vbNewLine & "よろしいですか？", vbYesNo) = vbNo Then Exit Sub
    Set s = ActiveSheet
    
    r = Selection.Item(1).row
    c = Selection.Item(1).column
    For Each n In ActiveWorkbook.names
        If InStr(n.RefersTo, "=" & s.Name & "!") = 1 _
                Or InStr(n.RefersTo, "='" & s.Name & "'!") = 1 Then
            s.Cells(r, c + 0).Value = "'" & n.Name
            s.Cells(r, c + 1).Value = "'" & n.RefersTo
            s.Cells(r, c + 2).Value = "'" & n.Comment
            r = r + 1
        End If
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のブックの名前定義をすべて削除し、選択中のセルの情報をもとに
'名前定義を作り直す。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180318   :sueki          :新規作成
'*******************************************************************************
Public Sub SA_名前定義刷新()
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '選択範囲確認
    If Selection.Columns.Count <> 2 Then
        MsgBox "選択範囲が不正です。" _
                & vbNewLine & "1列目に名称、2列目に定義範囲が入力されている範囲を選択してください。", vbExclamation
        Exit Sub
    End If
    
    '実行確認
    If MsgBox("現在設定されている名前定義をすべて破棄し、選択中のセル情報をもとに名前定義を作り直します。" _
                & vbNewLine & "よろしいですか？", vbYesNo) = vbNo Then Exit Sub
                
    For Each n In ActiveWorkbook.names
        n.Delete
    Next
    
    c = 1
    For r = 1 To Selection.Rows.Count
        If Selection.Item(r, c).Value <> "" _
                And Selection.Item(r, c + 1).Value <> "" Then
            ActiveWorkbook.names.Add Name:=Selection.Item(r, c).Value _
                                        , RefersTo:=Selection.Item(r, c + 1).Value
        End If
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のセルの情報をもとに名前定義を作成する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180318   :sueki          :新規作成
'*******************************************************************************
Public Sub SA_名前定義追加()
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '選択範囲確認
    If Selection.Columns.Count <> 2 Then
        MsgBox "選択範囲が不正です。" _
                & vbNewLine & "1列目に名称、2列目に定義範囲が入力されている範囲を選択してください。", vbExclamation
        Exit Sub
    End If
    
    '実行確認
    If MsgBox("選択中のセル情報をもとに名前定義を追加します。" _
                & vbNewLine & "すでに同じ名称の名前亭がある場合は上書きされます。" _
                & vbNewLine & "よろしいですか？", vbYesNo) = vbNo Then Exit Sub
    
    c = 1
    For r = 1 To Selection.Rows.Count
        If Selection.Item(r, c).Value <> "" _
                And Selection.Item(r, c + 1).Value <> "" Then
            ActiveWorkbook.names.Add Name:=Selection.Item(r, c).Value _
                                        , RefersTo:=Selection.Item(r, c + 1).Value
        End If
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のセルの情報をもとに名前定義を削除する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180318   :sueki          :新規作成
'*******************************************************************************
Public Sub SA_名前定義削除()
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '実行確認
    If MsgBox("選択中のセルの1列目に記載されている名前定義をすべて削除します。" _
                & vbNewLine & "よろしいですか？", vbYesNo) = vbNo Then Exit Sub
    
    c = 1
    For r = 1 To Selection.Rows.Count
        For Each n In ActiveWorkbook.names
            If n.Name = Selection.Item(r, 1).Value Then
                n.Delete
                Exit For
            End If
        Next
    Next
End Sub


