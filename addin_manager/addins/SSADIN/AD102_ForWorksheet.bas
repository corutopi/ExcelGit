Attribute VB_Name = "AD102_ForWorksheet"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'アクティブ状態のブックのすべての表示中シートの選択セルおよび表示
'状態をA1セルに整頓する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171230   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_ブック整頓()
    Dim ws As Worksheet
    Dim ti As Integer           'tmp index
    
    On Error GoTo ERR_MANAGER
    
    Call startProcess
    
    ti = 0
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If ti = 0 Then
                ti = ws.Index
            End If
            ws.Select
            Application.GoTo Range("A1"), True
        End If
    Next ws
    
    ActiveWorkbook.Worksheets(ti).Select
    
    Call endProcess
    Exit Sub
    
ERR_MANAGER:
    Call errProcess
End Sub


'*******************************************************************************
'-----------------------------------------
'アクティブ状態のブックのすべての非表示シートを表示状態にする。
'ブック保護による処理失敗には未対応。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171230   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_非表示シート全表示()
    Dim ws As Worksheet
    Dim ti As Integer           'tmp index
    
    On Error GoTo ERR_MANAGER
    
    Call startProcess
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
    Next ws
    
    Call endProcess
    Exit Sub
    
ERR_MANAGER:
    Call errProcess
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のブックのシート名を選択中のセルから下方向に列挙する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180119   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_シート名書出()
    Dim tb As Workbook
    Dim ts As Worksheet
    Dim ws As Worksheet
    Dim tr As Range
    
    Dim msg As String
    Dim r As Long
    Dim c As Long
    
    msg = "選択中のセルから下方向にアクティブブック内のシート名を列挙します。" _
            & vbNewLine & "実行してよろしいですか？"
            
    If MsgBox(msg, vbYesNo) = vbYes Then
        Set tb = ActiveWorkbook
        Set ts = ActiveSheet
        Set tr = Selection
        r = tr.Item(1).row
        c = tr.Item(1).column
        
        For Each ws In tb.Worksheets
            ts.Cells(r, c).Value = ws.Name
            r = r + 1
        Next
    End If
End Sub


'*******************************************************************************
'-----------------------------------------
'選択中のブックのシート名を変更する。
'選択中のセル範囲の1列目のシート名を2列目の名前に変更する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180119   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_シート名変更()
    Dim tb As Workbook
    Dim ts As Worksheet
    Dim ws As Worksheet
    Dim tr As Range
    
    Dim msg As String
    Dim r As Long
    Dim c As Long
    
    msg = "選択中のセル情報を元にシート名を変更します。" _
            & vbNewLine & "1列目のシート名を2列目の名前に変更します。" _
            & vbNewLine & "" _
            & vbNewLine & "よろしいですか？"
            
    If MsgBox(msg, vbYesNo) = vbYes Then
        Set tb = ActiveWorkbook
        Set ts = ActiveSheet
        Set tr = Selection
        
        For r = 1 To tr.Rows.Count
            For Each ws In ActiveWorkbook.Worksheets
                If tr.Item(r, 1).Value = ws.Name And tr.Item(r, 2).Value <> "" Then
                    ws.Name = tr.Item(r, 2).Value
                    Exit For
                End If
            Next
        Next
    End If
End Sub


'*******************************************************************************
'--SA_シート並び順変更-------------------------
'選択中のブックのシート名を変更する。
'選択中のセル範囲の1列目のシート名を2列目の名前に変更する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180205   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_シート並び順変更()
    Dim tb As Workbook
    Dim ts As Worksheet
    Dim ws As Worksheet
    Dim tr As Range
    
    Dim msg As String
    Dim r As Long
    Dim c As Long
    
    msg = "選択中のセル情報を元にシートの並び順を変更します。" _
            & vbNewLine & "1列目に記載のシートを上から順に先頭になるように並べます。" _
            & vbNewLine & "" _
            & vbNewLine & "よろしいですか？"
            
    If MsgBox(msg, vbYesNo) = vbYes Then
        Set tb = ActiveWorkbook
        Set ts = ActiveSheet
        Set tr = Selection
        
        For r = tr.Rows.Count To 1 Step -1
            For Each ws In ActiveWorkbook.Worksheets
                If tr.Item(r, 1).Value = ws.Name Then
                    ws.Move tb.Worksheets(1)
                    Exit For
                End If
            Next
        Next
    End If
End Sub



