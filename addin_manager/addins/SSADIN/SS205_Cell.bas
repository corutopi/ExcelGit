Attribute VB_Name = "SS205_Cell"
Option Explicit

'*******************************************************************************
'--comment_version_0.1.0------------------
'leftTop���N�_�Ƃ����ꍇ��target��Y���ʒu��Ԃ�
'leftTop�Ɠ����s�̏ꍇ��0�Ƃ���
'-----------------------------------------
'����       : *****
'�߂�l     : *****
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
Public Function RelativeR(target As Range, leftTop As Range) As Integer
    RelativeR = target.row - leftTop.row
End Function


'*******************************************************************************
'--comment_version_0.1.0------------------
'leftTop���N�_�Ƃ����ꍇ��target��X���ʒu��Ԃ�
'leftTop�Ɠ����s�̏ꍇ��0�Ƃ���
'-----------------------------------------
'����       : *****
'�߂�l     : *****
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
Public Function RelativeC(target As Range, leftTop As Range) As Integer
    RelativeC = target.column - leftTop.column
End Function


