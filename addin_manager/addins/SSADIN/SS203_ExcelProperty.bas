Attribute VB_Name = "SS203_ExcelProperty"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'ScreenUpdate�̏�Ԃ�ύX����
'-----------------------------------------
'����       :willBe         :�ύX��̏�Ԓl
'�߂�l     :�ύX�O�̏�Ԓl
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180701   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function ChangeScreenUpdate(willBe As Boolean) As Boolean
    ChangeScreenUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = willBe
End Function


'*******************************************************************************
'-----------------------------------------
'EnableEvents�̏�Ԃ�ύX����
'-----------------------------------------
'����       :willBe         :�ύX��̏�Ԓl
'�߂�l     :�ύX�O�̏�Ԓl
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180701   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function ChangeEnableEvents(willBe As Boolean) As Boolean
    ChangeEnableEvents = Application.EnableEvents
    Application.EnableEvents = willBe
End Function
