Attribute VB_Name = "AD101_ForWorkbook"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'���s���̃u�b�N�Ɠ����A�v���P�[�V�����ŊJ���Ă���G�N�Z�������ׂ�
'�ۑ������ɏI������B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180118   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�u�b�N�ꊇ�I��()
    Dim wb As Workbook
    
    If MsgBox("�����A�v���P�[�V�����ŊJ���Ă���G�N�Z�������ׂďI�����܂��B" _
                & vbNewLine & "(�ύX�͕ۑ�����܂���B)" _
                & vbNewLine _
                & vbNewLine & "���s���Ă�낵���ł����H", vbYesNo) = vbYes Then
        Application.DisplayAlerts = False
        Application.Quit
    End If
End Sub


'*******************************************************************************
'--SA_���W���[���C���|�[�g-------------------------
'�A�N�e�B�u��Ԃ̃u�b�N��(������)���W���[�����C���|�[�g����B
'���łɓ������W���[�������݂���ꍇ�͎����t�^���ꂽ�ʖ����W���[��
'�Ƃ��ăC���|�[�g����B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171230   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_���W���[���C���|�[�g()
    Dim wb          As Workbook
    Dim fdsi        As FileDialogSelectedItems
    Dim i, j        As Long
    
    On Error GoTo ERR_MANAGER
    
    Call startProcess

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "�C���|�[�g���郂�W���[���t�@�C���̑I��(�����I����)"
        .Filters.Add "���W���[���t�@�C��", "*.bas;*.cls;*.frm"
        If .Show = True Then
            Set fdsi = .SelectedItems
        Else
            GoTo END_MANAGER
        End If
    End With
    
    Set wb = ActiveWorkbook
    Debug.Print fdsi.Count
    
    For i = 1 To fdsi.Count
        Debug.Print fdsi(i)
        wb.VBProject.VBComponents.Import fdsi(i)
    Next i
    
    MsgBox "�C���|�[�g����"
    
END_MANAGER:
    Set wb = Nothing
    Set fdsi = Nothing
    Call endProcess
    Exit Sub
    
ERR_MANAGER:
    Call errProcess
End Sub


'*******************************************************************************
'--SA_���W���[���S�o��-------------------------
'�A�N�e�B�u��Ԃ̃u�b�N�̂��ׂẴ��W���[�����G�N�X�|�[�g����B
'�o�͂͑Ώۃu�b�N�Ɠ����t�H���_�ɍs���B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171230   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_���W���[���S�o��()
    Dim tm              As Object           'targetModule
    Dim outputPath      As String
    Dim fileExt         As String
    
    On Error GoTo ERR_MANAGER
    
    Call startProcess
    
    outputPath = ActiveWorkbook.path
    For Each tm In ActiveWorkbook.VBProject.VBComponents
        fileExt = GetExtFromModuleType(tm.Type)
        If fileExt <> "" Then
            Call ExportModule(tm, outputPath, fileExt)
        End If
    Next
    
    MsgBox "�o�͊���" _
            & vbNewLine & "�Ώۃu�b�N�Ɠ����t�H���_�ɏo�͂��܂����"
            
END_MANAGER:
    Call endProcess
    Exit Sub
    
ERR_MANAGER:
    Call errProcess
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'RC�Q�Ƃ�A1�Q�Ƃ�؂�ւ���.
'
'-----------------------------------------
'����       :����
'-----------------------------------------
'--�X�V����-------------------------------
'20181219   :xxx            :[�X�V���e]
'*******************************************************************************
Public Sub SA_RC�Q�Ɛؑ�()
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub

