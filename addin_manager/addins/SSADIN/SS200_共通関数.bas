Attribute VB_Name = "SS200_���ʊ֐�"
Option Explicit
Option Private Module


'*******************************************************************************
'--addBusinessDay-------------------------
'�Ώۓ��Ɏw�肵�������̉c�Ɠ��������Z����B
'
'-----------------------------------------
'����       :d              :�Ώۓ�
'����       :num            :���Z����
'�߂�l     :�v�Z����
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171220   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function AddBusinessDay(ByVal d As Date, ByVal num As Integer) As Date
    Dim s As Integer        '�X�e�b�v��
    Dim i As Integer
    
    If num >= 0 Then
        s = 1
    Else
        s = -1
    End If
    
    For i = s To num Step s
        d = d + s
        Do While Weekday(d, vbMonday) >= 6
            d = d + s
        Loop
    Next i
    
    AddBusinessDay = d
End Function


'*******************************************************************************
'--getFilepathFromDialog-------------------------
'�_�C�A���O����I�������t�@�C�����̔z����쐬����B
'-----------------------------------------
'����       :dialogTitle    :�I�v�V�����B�_�C�A���O�̃^�C�g���B
'����       :filterExt      :�I�v�V�����B�I���ł���t�@�C���̊g���q�`��������B"*.aaa;*.bbb;..."
'����       :filterTitle    :�I�v�V�����B�t�B���^�[�̃^�C�g��
'�߂�l     :�I���t�@�C���z��
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171230   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function GetFilepathFromDialog( _
        Optional ByVal dialogTitle As String = _
                "�Ώۃt�@�C���̑I��(�����I����)", _
        Optional ByVal filterExt As String = "", _
        Optional ByVal filterTitle As String = "�w�肳�ꂽ�g���q" _
        ) As Variant
    Dim re      As Variant
    Dim fdsi    As FileDialogSelectedItems
    Dim i       As Integer
    
    re = Null
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = dialogTitle
        .Filters.Clear
        If filterExt <> "" Then
            .Filters.Add filterTitle, filterExt
        End If
        If .Show = True Then
            Set fdsi = .SelectedItems
        Else
            GoTo END_MANAGER
        End If
    End With
    
    If fdsi.Count > 0 Then
        ReDim re(0)
        For i = 1 To fdsi.Count
            ReDim Preserve re(i - 1)
            re(i - 1) = fdsi.Item(i)
        Next
    End If
    
END_MANAGER:
    GetFilepathFromDialog = re
End Function


'*******************************************************************************
'--getFolderpathFromDialog-------------------------
'�_�C�A���O����I�������t�@�C�����̔z����쐬����B
'-----------------------------------------
'����       :dialogTitle    :�I�v�V�����B�_�C�A���O�̃^�C�g��
'�߂�l     :�I���t�H���_������
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171230   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function GetFolderpathFromDialog(Optional ByVal dialogTitle As String = "�Ώۃt�H���_�̑I��") As String
    Dim re      As String
    Dim fdsi    As FileDialogSelectedItems
    Dim i       As Integer
    
    re = ""
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = dialogTitle
        If .Show = True Then
            Set fdsi = .SelectedItems
        Else
            GoTo END_MANAGER
        End If
    End With
    
    If fdsi.Count > 0 Then
        re = fdsi.Item(1)
    End If
    
END_MANAGER:
    GetFolderpathFromDialog = re
End Function






