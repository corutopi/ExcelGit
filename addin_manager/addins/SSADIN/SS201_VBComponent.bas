Attribute VB_Name = "SS201_VBComponent"
'Microsoft Visual Basic for Applications Extensibilityの参照設定が必要

Option Explicit

'*******************************************************************************
'--exportAllModule-------------------------
'選択したファイル内のモジュールをすべてエクスポートする。
'-----------------------------------------
'引数       :対象ファイルパス
'引数       :出力先フォルダ。初期値は対象ファイルと同じフォルダ
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20190805   :xxx            :引数をファイルパスからWorkbookObjに変更
'20180701   :xxx            :新規作成
'*******************************************************************************
Public Sub ExportAllModule(wb As Workbook, Optional exportFolder = "")
    Dim tm              As Object           'targetModule
    Dim fileExt         As String
    
    If exportFolder = "" Then exportFolder = wb.path
    
    For Each tm In wb.VBProject.VBComponents
        fileExt = GetExtFromModuleType(tm.Type)
        If fileExt <> "" Then
            Call ExportModule(tm, exportFolder, fileExt)
        End If
    Next
    
    Exit Sub
End Sub

'*******************************************************************************
'--getExtFromModuleType-------------------------
'モジュールタイプ(VBComponent.Type)毎にテキストファイル出力時に
'使用する拡張子を取得する
'-----------------------------------------
'引数       :mType          :対象日
'戻り値     :対応する拡張子
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171220   :xxx            :新規作成
'*******************************************************************************
Public Function GetExtFromModuleType(ByVal mType As Integer) As String
    Dim re As String
    
    re = ""
    Select Case mType
        Case 1  'vbext_ct_StdModule
            re = "bas"
        Case 2, 100 'vbext_ct_ClassModule, vbext_ct_Document
            'ドキュメントモジュールはエクスポート対象から外してもいいかも
            re = "cls"
        Case 3  'vbext_ct_MSForm
            re = "frm"
    End Select

    GetExtFromModuleType = re
End Function

'*******************************************************************************
'--exportModule-------------------------
'モジュールを出力する
'-----------------------------------------
'引数       :target         :出力モジュールコンポーネント
'引数       :path           :出力先パス
'引数       :ext            :出力拡張子
'戻り値     :対応する拡張子
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171220   :xxx            :新規作成
'*******************************************************************************
Public Sub ExportModule(ByVal target As Object, ByVal path As String, ByVal ext As String)
    target.Export path & "\" & target.Name & "." & ext
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'モジュールをインポートする.
'既にインポートされているモジュールの場合は上書きするか別名でインポートする.
'-----------------------------------------
'引数       :hikihiki       :引数の説明
'戻り値     :戻り値の内容
'-----------------------------------------
'@todo バックアップするか否かもつけたら便利かも
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20190804   :               :新規作成
'*******************************************************************************
Public Sub ImportModule(modulePath As String, targetBook As Workbook, Optional overWrite As Boolean = True)
    '既存のモジュールかを確認
    If overWrite Then
        Dim moduleName As String
        Dim moduleExt As String
        Dim tm As Object  'VBComponent
        moduleName = Mid(modulePath, InStrRev(modulePath, "\") + 1)
        moduleName = Mid(moduleName, 1, InStrRev(moduleName, ".") - 1)  '拡張子の除去
        moduleExt = Mid(modulePath, InStrRev(modulePath, ".") + 1)
        For Each tm In targetBook.VBProject.VBComponents
            '同一モジュール判定
            If tm.Name <> moduleName Then
                '名称が一致していること -> ファイル名が変更されていないことが前提
            ElseIf GetExtFromModuleType(tm.Type) <> moduleExt Then
                '拡張子が一致すること
            ElseIf tm.Type = 100 Then
                'ドキュメントでないこと
            Else
                targetBook.VBProject.VBComponents.Remove tm
                Exit For
            End If
        Next
    End If
    'インポート
    targetBook.VBProject.VBComponents.Import modulePath
End Sub


