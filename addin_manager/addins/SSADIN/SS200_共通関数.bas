Attribute VB_Name = "SS200_共通関数"
Option Explicit
Option Private Module


'*******************************************************************************
'--addBusinessDay-------------------------
'対象日に指定した日数の営業日を加減算する。
'
'-----------------------------------------
'引数       :d              :対象日
'引数       :num            :加算日数
'戻り値     :計算結果
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171220   :xxx            :新規作成
'*******************************************************************************
Public Function AddBusinessDay(ByVal d As Date, ByVal num As Integer) As Date
    Dim s As Integer        'ステップ数
    Dim i As Integer
    
    If num >= 0 Then
        s = 1
    Else
        s = -1
    End If
    
    For i = s To num Step s
        d = d + s
        Do While Weekday(d, vbMonday) >= 6
            d = d + s
        Loop
    Next i
    
    AddBusinessDay = d
End Function


'*******************************************************************************
'--getFilepathFromDialog-------------------------
'ダイアログから選択したファイル名の配列を作成する。
'-----------------------------------------
'引数       :dialogTitle    :オプション。ダイアログのタイトル。
'引数       :filterExt      :オプション。選択できるファイルの拡張子形式文字列。"*.aaa;*.bbb;..."
'引数       :filterTitle    :オプション。フィルターのタイトル
'戻り値     :選択ファイル配列
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171230   :xxx            :新規作成
'*******************************************************************************
Public Function GetFilepathFromDialog( _
        Optional ByVal dialogTitle As String = _
                "対象ファイルの選択(複数選択可)", _
        Optional ByVal filterExt As String = "", _
        Optional ByVal filterTitle As String = "指定された拡張子" _
        ) As Variant
    Dim re      As Variant
    Dim fdsi    As FileDialogSelectedItems
    Dim i       As Integer
    
    re = Null
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = dialogTitle
        .Filters.Clear
        If filterExt <> "" Then
            .Filters.Add filterTitle, filterExt
        End If
        If .Show = True Then
            Set fdsi = .SelectedItems
        Else
            GoTo END_MANAGER
        End If
    End With
    
    If fdsi.Count > 0 Then
        ReDim re(0)
        For i = 1 To fdsi.Count
            ReDim Preserve re(i - 1)
            re(i - 1) = fdsi.Item(i)
        Next
    End If
    
END_MANAGER:
    GetFilepathFromDialog = re
End Function


'*******************************************************************************
'--getFolderpathFromDialog-------------------------
'ダイアログから選択したファイル名の配列を作成する。
'-----------------------------------------
'引数       :dialogTitle    :オプション。ダイアログのタイトル
'戻り値     :選択フォルダ文字列
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171230   :xxx            :新規作成
'*******************************************************************************
Public Function GetFolderpathFromDialog(Optional ByVal dialogTitle As String = "対象フォルダの選択") As String
    Dim re      As String
    Dim fdsi    As FileDialogSelectedItems
    Dim i       As Integer
    
    re = ""
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = dialogTitle
        If .Show = True Then
            Set fdsi = .SelectedItems
        Else
            GoTo END_MANAGER
        End If
    End With
    
    If fdsi.Count > 0 Then
        re = fdsi.Item(1)
    End If
    
END_MANAGER:
    GetFolderpathFromDialog = re
End Function






