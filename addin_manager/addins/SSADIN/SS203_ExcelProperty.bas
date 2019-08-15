Attribute VB_Name = "SS203_ExcelProperty"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'ScreenUpdateの状態を変更する
'-----------------------------------------
'引数       :willBe         :変更後の状態値
'戻り値     :変更前の状態値
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180701   :xxx            :新規作成
'*******************************************************************************
Public Function ChangeScreenUpdate(willBe As Boolean) As Boolean
    ChangeScreenUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = willBe
End Function


'*******************************************************************************
'-----------------------------------------
'EnableEventsの状態を変更する
'-----------------------------------------
'引数       :willBe         :変更後の状態値
'戻り値     :変更前の状態値
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'20180701   :xxx            :新規作成
'*******************************************************************************
Public Function ChangeEnableEvents(willBe As Boolean) As Boolean
    ChangeEnableEvents = Application.EnableEvents
    Application.EnableEvents = willBe
End Function
