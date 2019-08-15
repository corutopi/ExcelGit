Attribute VB_Name = "AD104_ForPicture"
Option Explicit

Private pr_height As Double
Private pr_width As Double

Private pr_trimTop As Double
Private pr_trimLeft As Double
Private pr_trimBottom As Double
Private pr_trimRight As Double


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃I�u�W�F�N�g(�܂��̓Z���͈�)�̃T�C�Y��ϐ��ɋL�^����B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171230   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�摜�T�C�Y�L��()
    On Error GoTo ERR_MANAGER
    
    pr_height = Selection.Height
    pr_width = Selection.Width
    
    Debug.Print "�����F" & pr_height
    Debug.Print "���@�F" & pr_width
    
    Exit Sub
    
ERR_MANAGER:
    Call errProcess("�T�C�Y���������擾�ł��܂���ł����B" _
            & vbNewLine & "�Z�����I�u�W�F�N�g��I��������ԂŎ��s���Ă��������B")
End Sub


'*******************************************************************************
'-----------------------------------------
'�uSA_�摜�T�C�Y�L���v�ŋL�������T�C�Y��I�𒆂̃I�u�W�F�N�g�ɔ��f����B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171230   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�摜�T�C�Y���f()
    On Error GoTo ERR_MANAGER
    
    If pr_height = 0 Or pr_width = 0 Then
        MsgBox "�摜�T�C�Y�L�����ŏ��ɍs���Ă��������B", vbInformation
    Else
        Selection.Height = pr_height
        Selection.Width = pr_width
    End If
    
    Exit Sub
    
ERR_MANAGER:
    Call errProcess("�T�C�Y�����������f����܂���ł���" _
            & vbNewLine & "�I�u�W�F�N�g��I��������ԂŎ��s���Ă��������B")
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃I�u�W�F�N�g�̃g���~���O�͈͂�ϐ��ɋL�^����B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180205   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�摜�g���~���O�͈͋L��()
    On Error GoTo ERR_MANAGER
    
    With Selection.PictureFormat
        pr_trimTop = .CropTop
        pr_trimLeft = .CropLeft
        pr_trimBottom = .CropBottom
        pr_trimRight = .CropRight
    End With
    
    Exit Sub
    
ERR_MANAGER:
    Call errProcess("�g���~���O�͈͂��������擾�ł��܂���ł����B" _
            & vbNewLine & "�I�u�W�F�N�g��I��������ԂŎ��s���Ă��������B")
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃I�u�W�F�N�g�ɋL�^�����g���~���O�͈͂𔽉f����B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180205   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�摜�g���~���O���f()
    On Error GoTo ERR_MANAGER
    
    With Selection.PictureFormat
        .CropTop = pr_trimTop
        .CropLeft = pr_trimLeft
        .CropBottom = pr_trimBottom
        .CropRight = pr_trimRight
    End With
    
    Exit Sub
    
ERR_MANAGER:
    Call errProcess("�T�C�Y�����������f����܂���ł���" _
            & vbNewLine & "�I�u�W�F�N�g��I��������ԂŎ��s���Ă��������B")
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̉摜���L������B
'
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180122   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�摜�_��()
    Const CL_AGREE_BEAST As String = "�_��b01"
    Dim ws As Worksheet
    
    Dim target As Shape
    Dim s As Shape
    
    On Error GoTo ERR_MANAGER
    
    If Selection.ShapeRange.Count > 1 Then GoTo ERR_MANAGER
    
    Set ws = ThisWorkbook.Worksheets(1)
    Set target = Selection.ShapeRange(1)
    
    For Each s In ws.Shapes
        If s.Name = CL_AGREE_BEAST Then
            s.Delete
            Exit For
        End If
    Next
    
    target.Copy
    ws.Paste
    
    Debug.Print Selection.Name
    
    ws.Shapes(ws.Shapes.Count).Name = CL_AGREE_BEAST
    
    Exit Sub
    
ERR_MANAGER:
    MsgBox "�_��Ɏ��s���܂����B" _
            & vbNewLine & "�P��̃I�u�W�F�N�g��I�����Ď��s���Ă��������B", vbExclamation
End Sub


'*******************************************************************************
'-----------------------------------------
'�_�񂵂��摜��\��t����B
'
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180122   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�_��摜����()
    Const CL_AGREE_BEAST As String = "�_��b01"
    
    On Error GoTo ERR_MANAGER
    ThisWorkbook.Worksheets(1).Shapes(CL_AGREE_BEAST).Copy
    ActiveSheet.Paste
    
    Exit Sub
    
ERR_MANAGER:
    MsgBox "�����Ɏ��s���܂����B" _
            & vbNewLine & "��ɉ摜�_����s���Ă��������B", vbExclamation
End Sub


