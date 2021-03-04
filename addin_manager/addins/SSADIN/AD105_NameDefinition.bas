Attribute VB_Name = "AD105_NameDefinition"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃u�b�N���̖��O��`��I�𒆂̃Z�����牺�����ɏ����o���B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20210305   :sueki          : add single quate before NameDefinition name
'                           : for Sheet names that inclued symbols.
'20180318   :sueki          : create
'*******************************************************************************
Public Sub SA_���O��`�����o��()
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '���s�m�F
    If MsgBox("�I�𒆂̃Z�����牺�����ɖ��O��`�̏����o�͂��܂�" _
                & vbNewLine & "��낵���ł����H", vbYesNo) = vbNo Then Exit Sub
    
    r = Selection.Item(1).row
    c = Selection.Item(1).column
    For Each n In ActiveWorkbook.names
        ActiveSheet.Cells(r, c + 0).Value = "'" & n.Name
        ActiveSheet.Cells(r, c + 1).Value = "'" & n.RefersTo
        ActiveSheet.Cells(r, c + 2).Value = "'" & n.Comment
        r = r + 1
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃V�[�g���̖��O��`��I�𒆂̃Z�����牺�����ɏ����o���B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20210305   :sueki          : add single quate before NameDefinition name
'                           : for Sheet names that inclued symbols.
'20180318   :sueki          : create
'*******************************************************************************
Public Sub SA_���O��`�����o��_�V�[�g�I�����[()
    Dim s As Worksheet
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '���s�m�F
    If MsgBox("�I�𒆂̃Z�����牺�����ɁA�I�𒆂̃V�[�g�ɎQ�Ɣ͈͂���`����Ă��閼�O��`�̏����o�͂��܂�" _
                & vbNewLine & "��낵���ł����H", vbYesNo) = vbNo Then Exit Sub
    Set s = ActiveSheet
    
    r = Selection.Item(1).row
    c = Selection.Item(1).column
    For Each n In ActiveWorkbook.names
        If InStr(n.RefersTo, "=" & s.Name & "!") = 1 _
                Or InStr(n.RefersTo, "='" & s.Name & "'!") = 1 Then
            s.Cells(r, c + 0).Value = "'" & n.Name
            s.Cells(r, c + 1).Value = "'" & n.RefersTo
            s.Cells(r, c + 2).Value = "'" & n.Comment
            r = r + 1
        End If
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃u�b�N�̖��O��`�����ׂč폜���A�I�𒆂̃Z���̏������Ƃ�
'���O��`����蒼���B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180318   :sueki          :�V�K�쐬
'*******************************************************************************
Public Sub SA_���O��`���V()
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '�I��͈͊m�F
    If Selection.Columns.Count <> 2 Then
        MsgBox "�I��͈͂��s���ł��B" _
                & vbNewLine & "1��ڂɖ��́A2��ڂɒ�`�͈͂����͂���Ă���͈͂�I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    '���s�m�F
    If MsgBox("���ݐݒ肳��Ă��閼�O��`�����ׂĔj�����A�I�𒆂̃Z���������Ƃɖ��O��`����蒼���܂��B" _
                & vbNewLine & "��낵���ł����H", vbYesNo) = vbNo Then Exit Sub
                
    For Each n In ActiveWorkbook.names
        n.Delete
    Next
    
    c = 1
    For r = 1 To Selection.Rows.Count
        If Selection.Item(r, c).Value <> "" _
                And Selection.Item(r, c + 1).Value <> "" Then
            ActiveWorkbook.names.Add Name:=Selection.Item(r, c).Value _
                                        , RefersTo:=Selection.Item(r, c + 1).Value
        End If
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃Z���̏������Ƃɖ��O��`���쐬����B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180318   :sueki          :�V�K�쐬
'*******************************************************************************
Public Sub SA_���O��`�ǉ�()
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '�I��͈͊m�F
    If Selection.Columns.Count <> 2 Then
        MsgBox "�I��͈͂��s���ł��B" _
                & vbNewLine & "1��ڂɖ��́A2��ڂɒ�`�͈͂����͂���Ă���͈͂�I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    '���s�m�F
    If MsgBox("�I�𒆂̃Z���������Ƃɖ��O��`��ǉ����܂��B" _
                & vbNewLine & "���łɓ������̖̂��O��������ꍇ�͏㏑������܂��B" _
                & vbNewLine & "��낵���ł����H", vbYesNo) = vbNo Then Exit Sub
    
    c = 1
    For r = 1 To Selection.Rows.Count
        If Selection.Item(r, c).Value <> "" _
                And Selection.Item(r, c + 1).Value <> "" Then
            ActiveWorkbook.names.Add Name:=Selection.Item(r, c).Value _
                                        , RefersTo:=Selection.Item(r, c + 1).Value
        End If
    Next
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃Z���̏������Ƃɖ��O��`���폜����B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180318   :sueki          :�V�K�쐬
'*******************************************************************************
Public Sub SA_���O��`�폜()
    Dim n As Name
    Dim r As Long
    Dim c As Long
    
    '���s�m�F
    If MsgBox("�I�𒆂̃Z����1��ڂɋL�ڂ���Ă��閼�O��`�����ׂč폜���܂��B" _
                & vbNewLine & "��낵���ł����H", vbYesNo) = vbNo Then Exit Sub
    
    c = 1
    For r = 1 To Selection.Rows.Count
        For Each n In ActiveWorkbook.names
            If n.Name = Selection.Item(r, 1).Value Then
                n.Delete
                Exit For
            End If
        Next
    Next
End Sub


