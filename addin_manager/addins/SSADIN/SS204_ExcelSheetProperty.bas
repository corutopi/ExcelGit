Attribute VB_Name = "SS204_ExcelSheetProperty"
Option Explicit


'*******************************************************************************
'--comment_version_0.1.0------------------
'�I�������Z���ɓ��͋K�������Ă�����Ă��邩�ǂ����𔻒肷��.
'�����͈̔͂̃Z�����w�肳��Ă���ꍇ��...�ǂ����悤.
'
'
'
'-----------------------------------------
'����       :hikihiki       :�����̐���
'�߂�l     :�߂�l�̓��e
'-----------------------------------------
'--�X�V����-------------------------------
'20181126   :xxx            :[�X�V���e]
'*******************************************************************************
Public Function HasVaridation(target As Range) As Boolean
    Const CP_TARGET_ERR As Long = 1004
    Dim re As Boolean
    Dim beforeErr As Long
    
    re = False
    beforeErr = Err.Number
    Err.Clear
    On Error Resume Next
    If target.Validation.Type Then
    End If
    If Err.Number = 0 Then
        re = True
    ElseIf Err.Number = CP_TARGET_ERR Then
        '�z��̃G���[�ł���Γ��͋K���Ȃ�
        re = False
    Else
        '���̑��̃G���[�̏ꍇ�͏����Ɏ��s�Ɣ��f
        On Error GoTo 0
        Err.Raise Err.Number
    End If
    If beforeErr <> 0 Then
        '�������{�O�ɃG���[���������Ă����ꍇ�͓����G���[������������Ԃɂ���
        Err.Raise beforeErr
    End If
    
    HasVaridation = re
End Function
