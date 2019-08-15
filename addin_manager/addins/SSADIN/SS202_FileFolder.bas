Attribute VB_Name = "SS202_FileFolder"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'フルパスからファイル名だけを切り出す
'
'-----------------------------------------
'引数       :fileFullPath   :ファイル名のフルパス
'戻り値     :ファイル名
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171220   :xxx            :新規作成
'*******************************************************************************
Public Function GetFilename(ByVal fileFullPath As String) As String
    Dim re As String
    
    re = Mid(fileFullPath, InStrRev(fileFullPath, "\") + 1)
    
    GetFilename = re
End Function


'*******************************************************************************
'-----------------------------------------
'ファイル名から拡張子だけを切り出す
'
'-----------------------------------------
'引数       :filename       :ファイル名
'戻り値     :拡張子
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171220   :xxx            :新規作成
'*******************************************************************************
Public Function GetExt(ByVal fileName As String) As String
    Dim re As String
    
    fileName = GetFilename(fileName)
    re = Mid(fileName, InStrRev(fileName, ".") + 1)
    
    If re = fileName Then
        re = ""
    End If
    
    GetExt = re
End Function


'*******************************************************************************
'-----------------------------------------
'フルパスからファイル名だけを切り出して拡張子を除く
'
'-----------------------------------------
'引数       :fileFullPath   :ファイル名の(フル)パス
'戻り値     :ファイル名
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20171220   :xxx            :新規作成
'*******************************************************************************
Public Function GetFilenameOnly(ByVal fileFullPath As String) As String
    Dim re As String
    
    re = Mid(fileFullPath, InStrRev(fileFullPath, "\") + 1)
    If InStr(re, ".") > 0 Then re = Left(re, InStr(re, ".") - 1)
    
    GetFilenameOnly = re
End Function


'*******************************************************************************
'-----------------------------------------
'指定されたパスの上位フォルダを返す。
'上位フォルダがない場合は引数をそのまま返す。
'-----------------------------------------
'引数       :path           :対象パス
'戻り値     :上位フォルダパス
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180701   :xxx            :新規作成
'*******************************************************************************
Public Function GetUpperFolder(ByVal path As String) As String
    Const CP_SEPPER As String = "\"
    Dim re          As String
    Dim sep         As Integer
    
    re = path
    
    If Right(re, 1) = CP_SEPPER Then
        re = Left(re, Len(re) - 1)
    End If
    
    sep = InStrRev(re, CP_SEPPER)
    If sep <> 0 Then
        re = Left(re, sep - 1)
    End If
    
    GetUpperFolder = re
End Function


'*******************************************************************************
'-----------------------------------------
'フォルダを上位パスを含めて作成する
'-----------------------------------------
'引数       :path           :作成対象パス
'戻り値     :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180701   :xxx            :新規作成
'*******************************************************************************
Public Sub MakeFolderRetroact(ByVal path As String)
    Dim fso As Variant
    Dim subPath As String
    
    On Error GoTo ERROR_MANAGER
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(path) = False Then
        subPath = GetUpperFolder(path)
        If subPath = path Then
            GoTo ERROR_MANAGER
        End If
        Call MakeFolderRetroact(subPath)
        fso.CreateFolder (path)
    End If
    
    Exit Sub
    
ERROR_MANAGER:
    Err.Raise 9999, , "フォルダ作成元となるドライブが確認できません。"
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'CSVの文字列データをCollection型の2重配列に変換する.
'ダブルクォーテーションで囲われたデータにも対応.
'全体の文字数によって処理時間が変わる.
'10 ** 7 文字で10秒程度.
'-----------------------------------------
'引数1      :csvData        :対象CSVファイルの全データ
'戻り値     :戻り値の内容
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20190815   :xxx            :新規作成
'*******************************************************************************
Public Function GetCSVCollection(csvData As String) As Collection
    Const DQ As String = """"
    Dim dataFlg As Boolean  'データ(文字列)の構築中にON
    Dim dqFlg As Boolean    'データがダブルクォーテーションに囲われている場合にON
    Dim tmpStr As String
    Dim ccol As Collection
    Dim rcol As Collection
    Dim i As Long
    Dim s As String
    Dim s2 As String
    
    dataFlg = True
    dqFlg = False
    
    i = 1
    Set ccol = New Collection
    Set rcol = New Collection
    Do While i <= Len(csvData) + 1
        s = Mid(csvData, i, 1)
        Select Case s
            Case DQ
                If dqFlg Then
                    s2 = Mid(csvData, i + 1, 1)
                    If s2 = DQ Then
                        i = i + 1
                        tmpStr = tmpStr & s
                    Else
                        ccol.Add tmpStr
                        dataFlg = False
                        dqFlg = False
                        tmpStr = ""
                    End If
                Else
                    If tmpStr = "" Then
                        dqFlg = True
                    Else
                        tmpStr = tmpStr & s
                    End If
                End If
            Case ","
                If dataFlg Then
                    If dqFlg Then
                        tmpStr = tmpStr & s
                    Else
                        ccol.Add tmpStr
                        dataFlg = False
                        dqFlg = False
                        '次のデータが始まる(のでフラグは結果的には立てたままになる)
                        dataFlg = True
                        tmpStr = ""
                    End If
                Else
                    dataFlg = True
                End If
            Case vbCr, vbLf, vbCrLf
                If dataFlg Then
                    If dqFlg Then
                        tmpStr = tmpStr & s
                    Else
                        '列データが更新される
                        If s = vbCr Then
                            If Mid(csvData, i + 1, 1) = vbLf Then
                                i = i + 1
                            End If
                        End If
                        'データの追加
                        ccol.Add tmpStr
                        dataFlg = False
                        dqFlg = False
                        '列の更新
                        rcol.Add ccol
                        dataFlg = True
                        Set ccol = New Collection
                        tmpStr = ""
                    End If
                Else
                    'DQ囲のデータ後にしか基本来ない想定
                    '列の更新
                    rcol.Add ccol
                    dataFlg = True
                    Set ccol = New Collection
                    tmpStr = ""
                End If
            Case Else
                If dataFlg Then
                    tmpStr = tmpStr & s
                End If
        End Select
        i = i + 1
    Loop
    Set GetCSVCollection = rcol
End Function
