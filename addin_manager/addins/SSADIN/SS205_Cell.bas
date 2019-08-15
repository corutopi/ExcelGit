Attribute VB_Name = "SS205_Cell"
Option Explicit

'*******************************************************************************
'--comment_version_0.1.0------------------
'leftTopを起点とした場合のtargetのY軸位置を返す
'leftTopと同じ行の場合を0とする
'-----------------------------------------
'引数       : *****
'戻り値     : *****
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
Public Function RelativeR(target As Range, leftTop As Range) As Integer
    RelativeR = target.row - leftTop.row
End Function


'*******************************************************************************
'--comment_version_0.1.0------------------
'leftTopを起点とした場合のtargetのX軸位置を返す
'leftTopと同じ行の場合を0とする
'-----------------------------------------
'引数       : *****
'戻り値     : *****
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************
Public Function RelativeC(target As Range, leftTop As Range) As Integer
    RelativeC = target.column - leftTop.column
End Function


