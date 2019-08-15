Attribute VB_Name = "AD103_ForCell"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃Z���ɑ΂����͂���Ă���l���ē��͂���B
'
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180122   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�Z���l�ē���()
    Dim r As Range
    
    Dim tmp As String
    Dim i As Long
    
    Set r = Selection
    
    For i = 1 To r.Count
        tmp = r.Item(i).Formula
        r.Item(i).Formula = tmp
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃Z����\�Ƃ��ăt�B���^�[�Ȃǂ��������C�A�E�g�ɉ��H����B
'
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20181219   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�\���C�A�E�g��()
    Const HEADER_COLOR = 16777185
    Dim target As Range
    
    '������
    Set target = Selection
    '���C�A�E�g���Z�b�g
    target.Borders.LineStyle = xlNone
    target.Font.Bold = False
    target.Interior.Pattern = xlNone
    '�r��
    With target.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    '�w�b�_�[�ҏW
    target.Rows(1).Interior.Color = HEADER_COLOR
    target.Rows(1).Font.Bold = True
    '�t�B���^�[�ݒ�
    If Not target.Parent.AutoFilter Is Nothing Then
        '�����̃t�B���^�[�̉���
        target.AutoFilter
    End If
    target.AutoFilter
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃Z���̃Z���������������Č����O�̒l�ŉ�����̃Z�����ׂĂ𖄂߂�B
'
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20181219   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_DB���f�[�^_�����[��()
    Dim target As Range
    Dim tmpMergeArea As Range
    Dim tmpValue As Variant
    Dim i As Long
    Dim j As Long
    
    Set target = Selection
    For i = 1 To target.Count
        If target.Item(i).MergeCells Then
            Set tmpMergeArea = target.Item(i).MergeArea
            tmpValue = target.Item(i).Value
            tmpMergeArea.UnMerge
            For j = 2 To tmpMergeArea.Count
                tmpMergeArea.Item(j).Value = tmpValue
            Next
        End If
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃Z���̋󔒃Z����1��̃Z���l�Ŗ��߂�
'
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20181219   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_DB���f�[�^_�󗓏[��()
    Dim target As Range
    Dim r As Long
    Dim c As Long
    
    Set target = Selection
    For c = 1 To target.Columns.Count
        For r = 2 To target.Rows.Count
            If target.Item(r, c).Value = "" Then
                target.Item(r, c).Value = target.Item(r - 1, c).Value
            End If
        Next
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃Z���͈͂̂��ׂĂ̍s���w�肵���s�����J����.
'�Z���͉������ɑ}���ړ�����.
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20181219   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�w��s���}��()
    Dim target As Range
    Dim tmpSU As Boolean
    Dim tmpStr As String
    Dim interval As Integer
    Dim i As Long
    
    '������
    tmpSU = ChangeScreenUpdate(False)
    '�}���񐔎擾
    Do
        tmpStr = InputBox("�}������s�����w�肵�Ă�������.")
        If tmpStr = "" Then
            Exit Sub
        ElseIf Not IsNumeric(tmpStr) Then
            MsgBox "�������w�肵�Ă�������.", vbExclamation
        ElseIf Not CDbl(tmpStr) = Int(tmpStr) Then
            MsgBox "�������w�肵�Ă�������.", vbExclamation
        Else
            interval = CInt(tmpStr)
            Exit Do
        End If
        tmpStr = ""
    Loop While tmpStr = ""
    '�f�[�^�}��
    Dim ws As Worksheet
    Dim insertRange As Range
    Dim tlRange As Range
    Dim brRange As Range
    
    Set ws = ActiveSheet
    Set target = Selection
    For i = 1 To target.Rows.Count - 1  '�ŏI�s�͑ΏۊO
        Set tlRange = ws.Cells(target.Item(1).row - 1 + 2 + (interval + 1) * (i - 1), _
                               target.Item(1).column)
        Set brRange = ws.Cells(target.Item(1).row - 1 + 2 + (interval + 1) * (i - 1) + (interval - 1), _
                               target.Item(target.Count).column)
        Set insertRange = ws.Range(tlRange, brRange)
        insertRange.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    '�I������
    ChangeScreenUpdate tmpSU
End Sub














