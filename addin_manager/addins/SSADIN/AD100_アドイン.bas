Attribute VB_Name = "AD100_アドイン"
Option Explicit


'*******************************************************************************
'--SA_えびぞう君起動-------------------------
'えびぞう君を起動する。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171230   :xxx            :新規作成
'*******************************************************************************
Public Sub SA_えびぞう君起動()
    SS701_Evidence.Show
End Sub


'*******************************************************************************
'--SA_フォルダ構成書出-------------------------
'選択したフォルダの構成情報をアクティブシートに書き出す。
'-----------------------------------------
'引数       :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180318   :sueki          :新規作成
'*******************************************************************************
Public Sub SA_フォルダ構成書出()
    
    Dim ws As Worksheet
    Dim parentPath As String
    Dim msg As String
    
    '確認ダイアログ
    msg = "選択中のシートを初期化してフォルダ構成の書き出しを行います。" _
            & vbNewLine & "よろしいですか？"
    
    If MsgBox(msg, vbYesNo) = vbNo Then Exit Sub
    
    'フォルダの選択
    parentPath = GetFolderpathFromDialog("親フォルダの選択")
    If parentPath = "" Then Exit Sub
    
    'シートの初期化
    Set ws = ActiveSheet
    ws.Cells.Clear
    
    'フォルダ構成書き出し開始
    Call makeFolderStructure(ws, parentPath)
    
    MsgBox "完了！"
End Sub
Private Function makeFolderStructure(ByVal ws As Worksheet _
                                        , ByVal path As String _
                                        , Optional ByVal r As Long = 0 _
                                        , Optional ByVal hierarchy As Integer = 0) As Long
    Const CL_MAX_HIERARCHY As Integer = 2
    
    Dim fso         As FileSystemObject
    Dim folder      As folder
    Dim file        As file
    Dim c As Long
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '最初の行(ヘッダー)のときのみの操作
    If hierarchy = 0 Then
        r = r + 1
        c = 0
        c = c + 1: ws.Cells(r, c).Value = "フルパス"
        c = c + 1: ws.Cells(r, c).Value = "所属階層"
        c = c + 1: ws.Cells(r, c).Value = "種別"
        For i = 1 To CL_MAX_HIERARCHY
            c = c + 1: ws.Cells(r, c).Value = "階層" & StrConv(Right("0" & i, 2), vbWide)
        Next
        c = c + 1: ws.Cells(r, c).Value = "サイズ(KB)"
        c = c + 1: ws.Cells(r, c).Value = "タイムスタンプ"
        
        '親フォルダ
        r = r + 1
        c = 0
        c = c + 1: ws.Cells(r, c).Value = path
        c = c + 1
        c = c + 1
        c = c + CL_MAX_HIERARCHY
        c = c + 1: ws.Cells(r, c).Value = "-"
        c = c + 1: ws.Cells(r, c).Value = "-"
        
        hierarchy = hierarchy + 1
    End If
    
    'フォルダの書出
    For Each folder In fso.GetFolder(path).SubFolders
        r = r + 1
        c = 0
        c = c + 1: ws.Cells(r, c).Value = folder.path
        c = c + 1: ws.Cells(r, c).Value = hierarchy
        c = c + 1: ws.Cells(r, c).Value = "フォルダ"
        
        c = c + hierarchy: ws.Cells(r, c).Value = folder.Name
        c = c + CL_MAX_HIERARCHY - hierarchy
        c = c + 1: ws.Cells(r, c).Value = folder.Size / 1000
        c = c + 1: ws.Cells(r, c).Value = folder.DateLastModified
        
        If hierarchy < CL_MAX_HIERARCHY Then
            r = makeFolderStructure(ws, folder.path, r, hierarchy + 1)
        End If
    Next
    
    'ファイルの書出
    For Each file In fso.GetFolder(path).Files
        r = r + 1
        c = 0
        c = c + 1: ws.Cells(r, c).Value = file.path
        c = c + 1: ws.Cells(r, c).Value = hierarchy
        c = c + 1: ws.Cells(r, c).Value = "ファイル"
        
        c = c + hierarchy: ws.Cells(r, c).Value = file.Name
        c = c + CL_MAX_HIERARCHY - hierarchy
        c = c + 1: ws.Cells(r, c).Value = file.Size / 1000
        c = c + 1: ws.Cells(r, c).Value = file.DateLastModified
    Next
    
    makeFolderStructure = r
End Function


'*******************************************************************************
'--comment_version_0.1.0------------------
'選択したセルから下方向に選択したフォルダ内のファイル名をすべて書き出す.
'
'-----------------------------------------
'引数       :無し
'-----------------------------------------
'--更新履歴-------------------------------
'20181219   :xxx            :[更新内容]
'*******************************************************************************
Public Sub SA_ファイル名書出()
    Dim folderPath As String
    Dim fileName As String
    Dim startCell As Range
    Dim i As Long
    '@todo ネットワーク越しのフォルダに対しては動かないかもしれない
    
    '対象フォルダ取得
    folderPath = GetFolderpathFromDialog()
    If folderPath = "" Then Exit Sub
    '起点セル
    Set startCell = Selection.Item(1)
    '出力元フォルダパス
    startCell.Value = folderPath
    'ファイル名全出力
    i = 1
    fileName = Dir(folderPath & "\*")
    Do While fileName <> ""
        startCell.Item(i + 1, 1).Value = fileName
        fileName = Dir()
        i = i + 1
    Loop
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'選択中のセルの内容をもとに指定したフォルダ内のファイル名を変更する.
'
'-----------------------------------------
'引数       :無し
'-----------------------------------------
'--更新履歴-------------------------------
'20181219   :xxx            :[更新内容]
'*******************************************************************************
Public Sub SA_ファイル名変更()
    Dim target As Range
    Dim folderPath As String
    Dim r As Long
    '@todo 正しく変換できるかどうかの確認処理を入れたほうが良いか？
    '@todo ネットワーク越しのフォルダに対しては動かないかもしれない
    
    '確認メッセージ
    MsgBox "選択中のセルの1列目のファイル名を2列目のファイル名に変更します。"
    '対象フォルダ取得
    folderPath = GetFolderpathFromDialog()
    If folderPath = "" Then Exit Sub
    Set target = Selection
    '名称変更
    For r = 1 To target.Rows.Count
        If target.Item(r, 1).Value <> "" And _
                folderPath & "\" & target.Item(r, 2).Value <> "" Then
            Name folderPath & "\" & target.Item(r, 1).Value As _
                    folderPath & "\" & target.Item(r, 2).Value
        End If
    Next
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'選択中のセルの内容をCSVデータとして出力する.
'
'-----------------------------------------
'引数       :無し
'-----------------------------------------
'--更新履歴-------------------------------
'20181219   :xxx            :[更新内容]
'*******************************************************************************
Public Sub SA_セルCSV出力()
    Dim target As Range
    Dim df As New SSC_MyDataFrame
    Dim filePath As String
    
    filePath = Application.GetSaveAsFilename( _
                        InitialFileName:="OutputCellData", _
                        FileFilter:="CSV File, *.csv")
    df.ReadCell Selection
    df.PrintData
    df.ExportCSV filePath
End Sub








