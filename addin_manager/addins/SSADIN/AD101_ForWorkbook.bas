Attribute VB_Name = "AD101_ForWorkbook"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'実行元のブックと同じアプリケーションで開いているエクセルをすべて
'保存せずに終了する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180118   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_ブック一括終了()
    Dim wb As Workbook
    
    If MsgBox("同じアプリケーションで開いているエクセルをすべて終了します。" _
                & vbNewLine & "(変更は保存されません。)" _
                & vbNewLine _
                & vbNewLine & "実行してよろしいですか？", vbYesNo) = vbYes Then
        Application.DisplayAlerts = False
        Application.Quit
    End If
End Sub


'*******************************************************************************
'--SA_モジュールインポート-------------------------
'アクティブ状態のブックに(複数の)モジュールをインポートする。
'すでに同名モジュールが存在する場合は自動付与された別名モジュール
'としてインポートする。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171230   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_モジュールインポート()
    Dim wb          As Workbook
    Dim fdsi        As FileDialogSelectedItems
    Dim i, j        As Long
    
    On Error GoTo ERR_MANAGER
    
    Call startProcess

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "インポートするモジュールファイルの選択(複数選択可)"
        .Filters.Add "モジュールファイル", "*.bas;*.cls;*.frm"
        If .Show = True Then
            Set fdsi = .SelectedItems
        Else
            GoTo END_MANAGER
        End If
    End With
    
    Set wb = ActiveWorkbook
    Debug.Print fdsi.Count
    
    For i = 1 To fdsi.Count
        Debug.Print fdsi(i)
        wb.VBProject.VBComponents.Import fdsi(i)
    Next i
    
    MsgBox "インポート完了"
    
END_MANAGER:
    Set wb = Nothing
    Set fdsi = Nothing
    Call endProcess
    Exit Sub
    
ERR_MANAGER:
    Call errProcess
End Sub


'*******************************************************************************
'--SA_モジュール全出力-------------------------
'アクティブ状態のブックのすべてのモジュールをエクスポートする。
'出力は対象ブックと同じフォルダに行う。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171230   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_モジュール全出力()
    Dim tm              As Object           'targetModule
    Dim outputPath      As String
    Dim fileExt         As String
    
    On Error GoTo ERR_MANAGER
    
    Call startProcess
    
    outputPath = ActiveWorkbook.path
    For Each tm In ActiveWorkbook.VBProject.VBComponents
        fileExt = GetExtFromModuleType(tm.Type)
        If fileExt <> "" Then
            Call ExportModule(tm, outputPath, fileExt)
        End If
    Next
    
    MsgBox "出力完了" _
            & vbNewLine & "対象ブックと同じフォルダに出力しました｡"
            
END_MANAGER:
    Call endProcess
    Exit Sub
    
ERR_MANAGER:
    Call errProcess
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'RC参照とA1参照を切り替える.
'
'-----------------------------------------
'引数       :無し
'-----------------------------------------
'--更新履歴-------------------------------
'20181219   :xxx            :[更新内容]
'*******************************************************************************
Public Sub SA_RC参照切替()
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub

