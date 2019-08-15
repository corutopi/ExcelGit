Attribute VB_Name = "SS501_TestAssist"
Option Explicit


'*******************************************************************************
'--comment_version_0.1.0------------------
'テストなどで使用するCSVファイルを作成する.
'実行時間は 100c * 10000r で10秒程度
'-----------------------------------------
'引数       :path           :出力先ファイルパス
'引数       :r              :データ行数
'引数       :c              :データ列数
'戻り値     :なし
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20190815   :xxx            :新規作成
'*******************************************************************************
Public Sub MakeRamdomCSV(path, r, c)
    Const STR_ARGS As String = "0123456789abcdef"
    Const DATA_LENGTH As Long = 16
    Dim fso As Object
    Dim data_r As String
    Dim data_c As String
    Dim tmpStr As String
    Dim argsLength As Long
    Dim i As Long, j As Long, k As Long
    Dim d As Date
    
    d = Now
    argsLength = Len(STR_ARGS)
    'path = "D:\Corutopi\040_Program\ExcelVBA\csv_test.csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.CreateTextFile(path)
        'ヘッダー
        tmpStr = ""
        For i = 1 To c
            tmpStr = tmpStr & ",data" & i
        Next
        .WriteLine Mid(tmpStr, 2)
        'データ
        For i = 1 To r
            data_r = ""
            For j = 1 To c
                data_c = ""
                For k = 1 To DATA_LENGTH
                    data_c = data_c & Mid(STR_ARGS, RndScope(1, argsLength), 1)
                Next
                data_r = data_r & "," & data_c
            Next
            .WriteLine Mid(data_r, 2)
        Next
    End With
    Debug.Print CDate(Now - d)
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'指定した範囲内のランダムな整数を取得する.
'どこかのサイトからもらってきたコード.
'-----------------------------------------
'引数       :MinNum         :乱数の最小値
'引数       :MaxNum         :乱数の最大値
'戻り値     :指定範囲内の任意の整数
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
Public Function RndScope(ByVal MinNum As Long, MaxNum As Long) As Long
    '指定した範囲の乱数を生成
    Dim ret As Integer
    Randomize
    ret = Int(Rnd() * (MaxNum - MinNum + 1) + MinNum)
    RndScope = ret
End Function
