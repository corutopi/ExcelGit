VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SSC_MyDataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'--comment_version_0.1.0------------------
'2次元データを扱いやすくするためのクラス.
'行列データを2重のCollectionクラスで保持し,各関数を使用してアクセスする。
'
'
'
'-----------------------------------------
'引数       :---            :
'戻り値     :---
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************

Option Explicit

'エラー関連
Private Const CP_ERR_NUMBER As Long = 9999
Private Const CP_ERR_MSG_FileNotFound As String = "ファイルが存在しません."
Private Const CP_ERR_MSG_CellAreaNum As String = "複数のセル範囲は指定できません."
Private Const CP_ERR_MSG_ArrayData As String = "引数に指定できるのはCollectionかオブジェクトを要素に含まないVariant配列のみです."
Private Const CP_ERR_MSG_DataColumnNum As String = "データ列数が一致しません."
Private Const CP_ERR_MSG_OptionArgs As String = "引数 {%s} は何れか {%n}つしか指定できません."
'クラス内列挙型
Private Enum EP_VALUE
    mdfNone
    mdfValueCollection
    mdfValueArray
    mdfValue
    mdfObject
End Enum

'要素
Private z_header As New Collection
Private z_datas As New Collection
Private z_index As New Collection  '検討中


'*******************************************************************************
'--comment_version_0.1.0------------------
'CSVファイルを読み込む
'
'-----------------------------------------
'引数       :path           :CSVファイルパス
'引数       :hasHeader      :ヘッダーの有無
'戻り値     :---
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
'CSVファイルを読み込む
Public Sub ReadCSV(path As String, Optional hasHeader As Boolean = True)
    Dim fso As Object
    Dim ts As Object
    Dim dataArr As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    'データチェック
    If Not fso.FileExists(path) Then
        Err.Raise Number:=CP_ERR_NUMBER, Description:=CP_ERR_MSG_FileNotFound
    End If
    '初期化
    Call Reset
    'ファイルオープン
    Set ts = fso.OpenTextFile(path, 1)  'read
    dataArr = Split(ts.ReadAll, vbCrLf)
    ts.Close
    '列数決定
    Dim i As Long
    Dim maxCol As Long
    Dim datas As Variant
    For i = LBound(dataArr) To UBound(dataArr)
        datas = Split(dataArr(i), ",")
        If maxCol <= UBound(datas) Then
            maxCol = UBound(datas) + 1
        End If
    Next
    'データ作成
    Dim tmpData As Collection
    Dim j As Long
    For i = LBound(dataArr) To UBound(dataArr)
        datas = Split(dataArr(i), ",")
        If i = 1 And hasHeader Then
            For j = LBound(datas) To UBound(datas)
                z_header.Add datas(j)
            Next
            For j = UBound(datas) + 1 To maxCol - 1
                z_header.Add ""
            Next
        Else
            Set tmpData = New Collection
            For j = LBound(datas) To UBound(datas)
                tmpData.Add datas(j)
            Next
            For j = UBound(datas) + 1 To maxCol - 1
                tmpData.Add ""
            Next
            z_datas.Add tmpData
        End If
    Next
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'セルデータを読み込む
'
'-----------------------------------------
'引数       :target         :読込セル範囲
'引数       :hasHeader      :ヘッダーの有無
'戻り値     :---
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
Public Sub ReadCell(target As Range, Optional hasHeader As Boolean = True)
    Dim r As Long
    Dim c As Long
    Dim tmpData As Collection
    
    'データチェック
    If target.Areas.Count > 1 Then
        Err.Raise Number:=CP_ERR_NUMBER, Description:=CP_ERR_MSG_CellAreaNum
    End If
    '初期化
    Call Reset
    'データ作成
    For r = 1 To target.Rows.Count
        If r = 1 And hasHeader Then
            For c = 1 To target.Columns.Count
                z_header.Add target.Item(r, c).Value
            Next
        Else
            Set tmpData = New Collection
            For c = 1 To target.Columns.Count
                tmpData.Add target.Item(r, c).Value
            Next
            z_datas.Add tmpData
        End If
    Next
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'CSVファイルに出力する
'
'-----------------------------------------
'引数       :path           :出力先ファイルパス
'戻り値     :---
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
Public Sub ExportCSV(path)
    Dim fso As Object
    Dim tmpStr As String
    Dim i As Long
    Dim j As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.CreateTextFile(path)
        'ヘッダーの出力
        tmpStr = ""
        If z_header.Count > 0 Then
            For i = 1 To Me.ColumnsCount
                tmpStr = tmpStr & "," & z_header.Item(i)
            Next
        End If
        .WriteLine Mid(tmpStr, 2)
        'データの出力
        For i = 1 To Me.Count
            tmpStr = ""
            For j = 1 To Me.ColumnsCount
                tmpStr = tmpStr & "," & z_datas.Item(i).Item(j)
            Next
            .WriteLine Mid(tmpStr, 2)
        Next
    End With
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'データをリセットする.
'保持しているヘッダー情報, データ情報をすべて破棄する.
'-----------------------------------------
'引数       :---            :
'戻り値     :---
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
Public Sub Reset()
    Set z_header = New Collection
    Set z_datas = New Collection
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'データの行数を取得する.
'
'-----------------------------------------
'引数       :---            :
'戻り値     :行数
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
Public Property Get Count()
    Count = z_datas.Count
End Property


'*******************************************************************************
'--comment_version_0.1.0------------------
'データの列数を取得する.
'
'-----------------------------------------
'引数       :---            :
'戻り値     :列数
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
Public Function ColumnsCount()
    Dim re As Long
    
    If z_header.Count <> 0 Then
        re = z_header.Count
    ElseIf z_datas.Count <> 0 Then
        re = z_datas.Item(1).Count
    Else
        re = 0
    End If
    ColumnsCount = re
End Function


'*******************************************************************************
'--comment_version_0.1.0------------------
'行データを追加する
'
'-----------------------------------------
'引数       :---            :
'戻り値     :列数
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
Public Sub AddRow(Data, Optional Before, Optional After)
    '@todo 自クラス型のデータが連携された場合の動作
    Dim tmpColl As New Collection
    Dim tmpColl2 As New Collection
    Dim i As Long
    
    'データのチェック
    Set tmpColl = ToCollection(Data)
    If tmpColl Is Nothing Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=CP_ERR_MSG_ArrayData
    End If
    If tmpColl.Count <> Me.Count Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=CP_ERR_MSG_DataColumnNum
    End If
    If IsMissing(Before) And IsMissing(After) Then
    ElseIf Not IsMissing(Before) And Not IsMissing(After) Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=CP_ERR_MSG_OptionArgs
    End If
    'データ追加(配列/コレクションの場合)
    For i = 1 To tmpColl.Count
        tmpColl2.Add tmpColl.Item(i), , Before, After
    Next
    z_datas.Add tmpColl
End Sub


Public Sub AddCol()

End Sub


'ヘッダーの設定
Public Sub SetHeader(header)
    Dim tmpColl As Collection
    
    'エラーチェック
    Set tmpColl = ToCollection(header)
    If tmpColl Is Nothing Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=""
    End If
    If Not (z_header.Count = 0 And z_datas.Count = 0) And _
            ColumnsCount <> tmpColl.Count Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=CP_ERR_MSG_DataColumnNum
    End If
    'ヘッダーの再設定
    Set z_header = tmpColl  'ToCollectionで新しいオブジェクトになってるからOK
End Sub


'データの抽出
Public Function GetValue(row, column)
    Dim tmpRow As Long
    Dim tmpColumn As Long
    
    tmpRow = row
    If column = Int(column) Then
        tmpColumn = column
    Else
        tmpColumn = GetColumnNum(column)
    End If
    GetValue = z_datas.Itemm(tmpRow).Item(tmpColumn)
End Function


'データの更新
Public Sub SetValue(row, column, val)
    Dim tmpRow As Long
    Dim tmpColumn As Long
    
    tmpRow = row
    If column = Int(column) Then
        tmpColumn = column
    Else
        tmpColumn = GetColumnNum(column)
    End If
    z_datas.Itemm(tmpRow).Add val, Before:=column
    z_datas.Itemm(tmpRow).Remove column + 1
End Sub


'データ行の削除
Public Sub Delete(row)
    Dim tmpRow As Long
    
    tmpRow = row
    z_datas.Remove tmpRow
End Sub


'ヘッダーの名称⇔番号変換
Public Function GetColumnNum(Name) As Long
    Dim re As Long
    Dim i As Long
    
    re = -1
    For i = 1 To z_header.Count
        If z_header.Item(i) = Name Then
            re = i
            Exit For
        End If
    Next
    GetColumnNum = re
End Function


'データの連結（行）


'データのソート
Public Function Sort(Columns, Optional IsAsc As Boolean = True) As SSC_MyDataFrame
    Dim re As New SSC_MyDataFrame
    Dim tmpColl As New Collection
    Dim sortTarget As Variant
    Dim i As Long
    Dim j As Long
    
    Set tmpColl = ToCollection(Columns)
    sortTarget = tmpColl.Item(tmpColl.Count) '最後のデータでソートする
    For i = 1 To Me.Count
        Dim beforePoint As Long
        beforePoint = -1
        For j = 1 To re.Count
            If (Me.GetValue(i, sortTarget) < re.GetValue(j, sortTarget)) = IsAsc Then
                beforePoint = j
                Exit For
            End If
        Next
        If beforePoint = -1 Then
            '1行丸ごとのデータをどう取得するか？
            re.AddRow z_datas.Item(j)
        Else
            re.AddRow z_datas.Item(j), Before:=beforePoint
        End If
    Next
    
    If tmpColl.Count > 1 Then
        tmpColl.Remove tmpColl.Count
        Set re = re.Sort(tmpColl, IsAsc)
    End If
    Set Sort = re
End Function


'データの一意化
Public Function UniqueData(Optional Columns) As SSC_MyDataFrame
    Dim tmpColl As New Collection
    Dim columnNums As New Collection
    Dim re As SSC_MyDataFrame
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    '確認対象列の決定
    If IsMissing(Columns) Then
        For i = 1 To Me.Count
            columnNums.Add i
        Next
    Else
        Set tmpColl = ToCollection(Columns)
        For i = 1 To tmpColl.Count
            columnNums.Add Me.GetColumnNum(tmpColl.Item(i))
        Next
    End If
    
    '重複データの削除
    Set re = Me.Copy()
    i = 1
    Do While i <= re.Count
        Dim check As Boolean
        check = True  '確認対象列がすべて一致するデータがあればTrue
        For j = i - 1 To 1 Step -1
            For k = 1 To columnNums.Count
                If re.GetValue(i, k) <> re.GetValue(j, k) Then
                    check = False
                    Exit For
                End If
            Next
            If check = False Then Exit For
        Next
        If check Then
            re.Delete i
        Else
            i = i + 1
        End If
    Loop
    Set UniqueData = re
End Function


'データの選択
Public Function SelectData(column, Value, Optional isEqual = True) As SSC_MyDataFrame
    Dim columnNum As Long
    Dim re As SSC_MyDataFrame
    Dim i As Long
    
    Set re = Me.Copy()
    i = 1
    Do While i <= re.Count
        If re.GetValue(i, column) <> Value Then
            re.Delete i
        End If
    Loop
    Set SelectData = re
End Function


'データをコピーする
Public Function Copy() As SSC_MyDataFrame
    Dim re As New SSC_MyDataFrame
    Dim i As Long
    
    If z_header.Count > 0 Then
        re.SetHeader ToCollection(z_header)
    End If
    For i = 1 To Me.Count
        re.AddRow ToCollection(z_datas.Item(i))
    Next
    
    Set Copy = re
End Function


'データを出力する
Public Sub PrintData()
    Dim maxChara As New Collection  '行ごとの最大出力文字数
    Dim r As Long
    Dim c As Long
    Dim tmpChara As Long
    Dim tmpStr As String
    
    '最大文字数決定
    For c = 1 To Me.ColumnsCount
        tmpChara = 0
        If z_header.Count > 0 Then
            tmpChara = GetByteCharaNum(z_header.Item(c))
        End If
        For r = 1 To z_datas.Count
            If tmpChara < GetByteCharaNum(z_datas.Item(r).Item(c)) Then
                tmpChara = GetByteCharaNum(z_datas.Item(r).Item(c))
            End If
        Next
        maxChara.Add tmpChara
    Next
    'データ出力
    tmpStr = ""
    If z_header.Count > 0 Then
        For c = 1 To z_header.Count
            tmpStr = tmpStr & "," & String(maxChara.Item(c) - GetByteCharaNum(z_header.Item(c)), " ") & z_header.Item(c)
        Next
        Debug.Print Mid(tmpStr, 2)
    Else
        Debug.Print "[None Header]"
    End If
    For r = 1 To z_datas.Count
        tmpStr = ""
        For c = 1 To z_datas.Item(r).Count
            tmpStr = tmpStr & "," & String(maxChara.Item(c) - GetByteCharaNum(z_datas.Item(r).Item(c)), " ") & z_datas.Item(r).Item(c)
        Next
        Debug.Print Mid(tmpStr, 2)
    Next
    If z_datas.Count = 0 Then
        Debug.Print "[None Data]"
    End If
End Sub


'*******************************************************************************
'ここからクラス内メソッド*******************************************************
'*******************************************************************************
'バイト文字数を取得する
Private Function GetByteCharaNum(str) As Long
    GetByteCharaNum = LenB(StrConv(str, vbFromUnicode))
End Function


'配列/単一の値をコレクション型に変換する
Private Function ToCollection(val) As Collection
    Dim re As New Collection
    Dim valType As EP_VALUE
    Dim i As Long
    
    valType = GetValueType(val)
    
    If valType = mdfValue Then
        re.Add val
    ElseIf valType = mdfValueArray Then
        For i = LBound(val) To UBound(val)
            re.Add val(i)
        Next
    ElseIf valType = mdfValueCollection Then
        For i = 1 To val.Count
            re.Add val.Item(i)
        Next
    Else
        Set re = Nothing
    End If
    Set ToCollection = re
End Function


'引数がこのクラスで使用可能な配列/コレクション/単一の値であることを確認する
Private Function GetValueType(obj) As EP_VALUE
    Dim re As EP_VALUE
    Dim arrCheck As EP_VALUE
    Dim i As Long
    re = mdfNone
    If TypeName(obj) = "Collection" Then
        'コレクション
        arrCheck = mdfValueCollection
        For i = 1 To obj.Count
            If IsObject(obj.Item(i)) Then
                arrCheck = mdfObject
                Exit For
            End If
        Next
        re = arrCheck
    ElseIf IsArray(obj) Then
        '配列
        arrCheck = mdfValueArray
        For i = LBound(obj) To UBound(obj)
            If IsObject(obj(i)) Then
                arrCheck = mdfObject
                Exit For
            End If
        Next
        re = arrCheck
    ElseIf IsObject(obj) Then
        '単一の値
        re = mdfValue
    Else
        'そのほかの値は不可
        re = mdfObject
    End If
    
    GetValueType = re
End Function


'エラーメッセージを作成する @todo 検討中
Private Function MakeErrMsg(msg, s, n) As String

End Function

