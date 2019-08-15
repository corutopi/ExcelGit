Attribute VB_Name = "SS202_FileFolder"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'�t���p�X����t�@�C����������؂�o��
'
'-----------------------------------------
'����       :fileFullPath   :�t�@�C�����̃t���p�X
'�߂�l     :�t�@�C����
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171220   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function GetFilename(ByVal fileFullPath As String) As String
    Dim re As String
    
    re = Mid(fileFullPath, InStrRev(fileFullPath, "\") + 1)
    
    GetFilename = re
End Function


'*******************************************************************************
'-----------------------------------------
'�t�@�C��������g���q������؂�o��
'
'-----------------------------------------
'����       :filename       :�t�@�C����
'�߂�l     :�g���q
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171220   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function GetExt(ByVal fileName As String) As String
    Dim re As String
    
    fileName = GetFilename(fileName)
    re = Mid(fileName, InStrRev(fileName, ".") + 1)
    
    If re = fileName Then
        re = ""
    End If
    
    GetExt = re
End Function


'*******************************************************************************
'-----------------------------------------
'�t���p�X����t�@�C����������؂�o���Ċg���q������
'
'-----------------------------------------
'����       :fileFullPath   :�t�@�C������(�t��)�p�X
'�߂�l     :�t�@�C����
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171220   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function GetFilenameOnly(ByVal fileFullPath As String) As String
    Dim re As String
    
    re = Mid(fileFullPath, InStrRev(fileFullPath, "\") + 1)
    If InStr(re, ".") > 0 Then re = Left(re, InStr(re, ".") - 1)
    
    GetFilenameOnly = re
End Function


'*******************************************************************************
'-----------------------------------------
'�w�肳�ꂽ�p�X�̏�ʃt�H���_��Ԃ��B
'��ʃt�H���_���Ȃ��ꍇ�͈��������̂܂ܕԂ��B
'-----------------------------------------
'����       :path           :�Ώۃp�X
'�߂�l     :��ʃt�H���_�p�X
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180701   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function GetUpperFolder(ByVal path As String) As String
    Const CP_SEPPER As String = "\"
    Dim re          As String
    Dim sep         As Integer
    
    re = path
    
    If Right(re, 1) = CP_SEPPER Then
        re = Left(re, Len(re) - 1)
    End If
    
    sep = InStrRev(re, CP_SEPPER)
    If sep <> 0 Then
        re = Left(re, sep - 1)
    End If
    
    GetUpperFolder = re
End Function


'*******************************************************************************
'-----------------------------------------
'�t�H���_����ʃp�X���܂߂č쐬����
'-----------------------------------------
'����       :path           :�쐬�Ώۃp�X
'�߂�l     :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180701   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub MakeFolderRetroact(ByVal path As String)
    Dim fso As Variant
    Dim subPath As String
    
    On Error GoTo ERROR_MANAGER
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(path) = False Then
        subPath = GetUpperFolder(path)
        If subPath = path Then
            GoTo ERROR_MANAGER
        End If
        Call MakeFolderRetroact(subPath)
        fso.CreateFolder (path)
    End If
    
    Exit Sub
    
ERROR_MANAGER:
    Err.Raise 9999, , "�t�H���_�쐬���ƂȂ�h���C�u���m�F�ł��܂���B"
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'CSV�̕�����f�[�^��Collection�^��2�d�z��ɕϊ�����.
'�_�u���N�H�[�e�[�V�����ň͂�ꂽ�f�[�^�ɂ��Ή�.
'�S�̂̕������ɂ���ď������Ԃ��ς��.
'10 ** 7 ������10�b���x.
'-----------------------------------------
'����1      :csvData        :�Ώ�CSV�t�@�C���̑S�f�[�^
'�߂�l     :�߂�l�̓��e
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20190815   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function GetCSVCollection(csvData As String) As Collection
    Const DQ As String = """"
    Dim dataFlg As Boolean  '�f�[�^(������)�̍\�z����ON
    Dim dqFlg As Boolean    '�f�[�^���_�u���N�H�[�e�[�V�����Ɉ͂��Ă���ꍇ��ON
    Dim tmpStr As String
    Dim ccol As Collection
    Dim rcol As Collection
    Dim i As Long
    Dim s As String
    Dim s2 As String
    
    dataFlg = True
    dqFlg = False
    
    i = 1
    Set ccol = New Collection
    Set rcol = New Collection
    Do While i <= Len(csvData) + 1
        s = Mid(csvData, i, 1)
        Select Case s
            Case DQ
                If dqFlg Then
                    s2 = Mid(csvData, i + 1, 1)
                    If s2 = DQ Then
                        i = i + 1
                        tmpStr = tmpStr & s
                    Else
                        ccol.Add tmpStr
                        dataFlg = False
                        dqFlg = False
                        tmpStr = ""
                    End If
                Else
                    If tmpStr = "" Then
                        dqFlg = True
                    Else
                        tmpStr = tmpStr & s
                    End If
                End If
            Case ","
                If dataFlg Then
                    If dqFlg Then
                        tmpStr = tmpStr & s
                    Else
                        ccol.Add tmpStr
                        dataFlg = False
                        dqFlg = False
                        '���̃f�[�^���n�܂�(�̂Ńt���O�͌��ʓI�ɂ͗��Ă��܂܂ɂȂ�)
                        dataFlg = True
                        tmpStr = ""
                    End If
                Else
                    dataFlg = True
                End If
            Case vbCr, vbLf, vbCrLf
                If dataFlg Then
                    If dqFlg Then
                        tmpStr = tmpStr & s
                    Else
                        '��f�[�^���X�V�����
                        If s = vbCr Then
                            If Mid(csvData, i + 1, 1) = vbLf Then
                                i = i + 1
                            End If
                        End If
                        '�f�[�^�̒ǉ�
                        ccol.Add tmpStr
                        dataFlg = False
                        dqFlg = False
                        '��̍X�V
                        rcol.Add ccol
                        dataFlg = True
                        Set ccol = New Collection
                        tmpStr = ""
                    End If
                Else
                    'DQ�͂̃f�[�^��ɂ�����{���Ȃ��z��
                    '��̍X�V
                    rcol.Add ccol
                    dataFlg = True
                    Set ccol = New Collection
                    tmpStr = ""
                End If
            Case Else
                If dataFlg Then
                    tmpStr = tmpStr & s
                End If
        End Select
        i = i + 1
    Loop
    Set GetCSVCollection = rcol
End Function
