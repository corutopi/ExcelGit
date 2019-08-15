Attribute VB_Name = "AD100_�A�h�C��"
Option Explicit


'*******************************************************************************
'--SA_���т����N�N��-------------------------
'���т����N���N������B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171230   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_���т����N�N��()
    SS701_Evidence.Show
End Sub


'*******************************************************************************
'--SA_�t�H���_�\�����o-------------------------
'�I�������t�H���_�̍\�������A�N�e�B�u�V�[�g�ɏ����o���B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180318   :sueki          :�V�K�쐬
'*******************************************************************************
Public Sub SA_�t�H���_�\�����o()
    
    Dim ws As Worksheet
    Dim parentPath As String
    Dim msg As String
    
    '�m�F�_�C�A���O
    msg = "�I�𒆂̃V�[�g�����������ăt�H���_�\���̏����o�����s���܂��B" _
            & vbNewLine & "��낵���ł����H"
    
    If MsgBox(msg, vbYesNo) = vbNo Then Exit Sub
    
    '�t�H���_�̑I��
    parentPath = GetFolderpathFromDialog("�e�t�H���_�̑I��")
    If parentPath = "" Then Exit Sub
    
    '�V�[�g�̏�����
    Set ws = ActiveSheet
    ws.Cells.Clear
    
    '�t�H���_�\�������o���J�n
    Call makeFolderStructure(ws, parentPath)
    
    MsgBox "�����I"
End Sub
Private Function makeFolderStructure(ByVal ws As Worksheet _
                                        , ByVal path As String _
                                        , Optional ByVal r As Long = 0 _
                                        , Optional ByVal hierarchy As Integer = 0) As Long
    Const CL_MAX_HIERARCHY As Integer = 2
    
    Dim fso         As FileSystemObject
    Dim folder      As folder
    Dim file        As file
    Dim c As Long
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�ŏ��̍s(�w�b�_�[)�̂Ƃ��݂̂̑���
    If hierarchy = 0 Then
        r = r + 1
        c = 0
        c = c + 1: ws.Cells(r, c).Value = "�t���p�X"
        c = c + 1: ws.Cells(r, c).Value = "�����K�w"
        c = c + 1: ws.Cells(r, c).Value = "���"
        For i = 1 To CL_MAX_HIERARCHY
            c = c + 1: ws.Cells(r, c).Value = "�K�w" & StrConv(Right("0" & i, 2), vbWide)
        Next
        c = c + 1: ws.Cells(r, c).Value = "�T�C�Y(KB)"
        c = c + 1: ws.Cells(r, c).Value = "�^�C���X�^���v"
        
        '�e�t�H���_
        r = r + 1
        c = 0
        c = c + 1: ws.Cells(r, c).Value = path
        c = c + 1
        c = c + 1
        c = c + CL_MAX_HIERARCHY
        c = c + 1: ws.Cells(r, c).Value = "-"
        c = c + 1: ws.Cells(r, c).Value = "-"
        
        hierarchy = hierarchy + 1
    End If
    
    '�t�H���_�̏��o
    For Each folder In fso.GetFolder(path).SubFolders
        r = r + 1
        c = 0
        c = c + 1: ws.Cells(r, c).Value = folder.path
        c = c + 1: ws.Cells(r, c).Value = hierarchy
        c = c + 1: ws.Cells(r, c).Value = "�t�H���_"
        
        c = c + hierarchy: ws.Cells(r, c).Value = folder.Name
        c = c + CL_MAX_HIERARCHY - hierarchy
        c = c + 1: ws.Cells(r, c).Value = folder.Size / 1000
        c = c + 1: ws.Cells(r, c).Value = folder.DateLastModified
        
        If hierarchy < CL_MAX_HIERARCHY Then
            r = makeFolderStructure(ws, folder.path, r, hierarchy + 1)
        End If
    Next
    
    '�t�@�C���̏��o
    For Each file In fso.GetFolder(path).Files
        r = r + 1
        c = 0
        c = c + 1: ws.Cells(r, c).Value = file.path
        c = c + 1: ws.Cells(r, c).Value = hierarchy
        c = c + 1: ws.Cells(r, c).Value = "�t�@�C��"
        
        c = c + hierarchy: ws.Cells(r, c).Value = file.Name
        c = c + CL_MAX_HIERARCHY - hierarchy
        c = c + 1: ws.Cells(r, c).Value = file.Size / 1000
        c = c + 1: ws.Cells(r, c).Value = file.DateLastModified
    Next
    
    makeFolderStructure = r
End Function


'*******************************************************************************
'--comment_version_0.1.0------------------
'�I�������Z�����牺�����ɑI�������t�H���_���̃t�@�C���������ׂď����o��.
'
'-----------------------------------------
'����       :����
'-----------------------------------------
'--�X�V����-------------------------------
'20181219   :xxx            :[�X�V���e]
'*******************************************************************************
Public Sub SA_�t�@�C�������o()
    Dim folderPath As String
    Dim fileName As String
    Dim startCell As Range
    Dim i As Long
    '@todo �l�b�g���[�N�z���̃t�H���_�ɑ΂��Ă͓����Ȃ���������Ȃ�
    
    '�Ώۃt�H���_�擾
    folderPath = GetFolderpathFromDialog()
    If folderPath = "" Then Exit Sub
    '�N�_�Z��
    Set startCell = Selection.Item(1)
    '�o�͌��t�H���_�p�X
    startCell.Value = folderPath
    '�t�@�C�����S�o��
    i = 1
    fileName = Dir(folderPath & "\*")
    Do While fileName <> ""
        startCell.Item(i + 1, 1).Value = fileName
        fileName = Dir()
        i = i + 1
    Loop
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'�I�𒆂̃Z���̓��e�����ƂɎw�肵���t�H���_���̃t�@�C������ύX����.
'
'-----------------------------------------
'����       :����
'-----------------------------------------
'--�X�V����-------------------------------
'20181219   :xxx            :[�X�V���e]
'*******************************************************************************
Public Sub SA_�t�@�C�����ύX()
    Dim target As Range
    Dim folderPath As String
    Dim r As Long
    '@todo �������ϊ��ł��邩�ǂ����̊m�F��������ꂽ�ق����ǂ����H
    '@todo �l�b�g���[�N�z���̃t�H���_�ɑ΂��Ă͓����Ȃ���������Ȃ�
    
    '�m�F���b�Z�[�W
    MsgBox "�I�𒆂̃Z����1��ڂ̃t�@�C������2��ڂ̃t�@�C�����ɕύX���܂��B"
    '�Ώۃt�H���_�擾
    folderPath = GetFolderpathFromDialog()
    If folderPath = "" Then Exit Sub
    Set target = Selection
    '���̕ύX
    For r = 1 To target.Rows.Count
        If target.Item(r, 1).Value <> "" And _
                folderPath & "\" & target.Item(r, 2).Value <> "" Then
            Name folderPath & "\" & target.Item(r, 1).Value As _
                    folderPath & "\" & target.Item(r, 2).Value
        End If
    Next
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'�I�𒆂̃Z���̓��e��CSV�f�[�^�Ƃ��ďo�͂���.
'
'-----------------------------------------
'����       :����
'-----------------------------------------
'--�X�V����-------------------------------
'20181219   :xxx            :[�X�V���e]
'*******************************************************************************
Public Sub SA_�Z��CSV�o��()
    Dim target As Range
    Dim df As New SSC_MyDataFrame
    Dim filePath As String
    
    filePath = Application.GetSaveAsFilename( _
                        InitialFileName:="OutputCellData", _
                        FileFilter:="CSV File, *.csv")
    df.ReadCell Selection
    df.PrintData
    df.ExportCSV filePath
End Sub








