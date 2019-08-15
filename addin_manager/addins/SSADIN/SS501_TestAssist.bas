Attribute VB_Name = "SS501_TestAssist"
Option Explicit


'*******************************************************************************
'--comment_version_0.1.0------------------
'�e�X�g�ȂǂŎg�p����CSV�t�@�C�����쐬����.
'���s���Ԃ� 100c * 10000r ��10�b���x
'-----------------------------------------
'����       :path           :�o�͐�t�@�C���p�X
'����       :r              :�f�[�^�s��
'����       :c              :�f�[�^��
'�߂�l     :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20190815   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub MakeRamdomCSV(path, r, c)
    Const STR_ARGS As String = "0123456789abcdef"
    Const DATA_LENGTH As Long = 16
    Dim fso As Object
    Dim data_r As String
    Dim data_c As String
    Dim tmpStr As String
    Dim argsLength As Long
    Dim i As Long, j As Long, k As Long
    Dim d As Date
    
    d = Now
    argsLength = Len(STR_ARGS)
    'path = "D:\Corutopi\040_Program\ExcelVBA\csv_test.csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.CreateTextFile(path)
        '�w�b�_�[
        tmpStr = ""
        For i = 1 To c
            tmpStr = tmpStr & ",data" & i
        Next
        .WriteLine Mid(tmpStr, 2)
        '�f�[�^
        For i = 1 To r
            data_r = ""
            For j = 1 To c
                data_c = ""
                For k = 1 To DATA_LENGTH
                    data_c = data_c & Mid(STR_ARGS, RndScope(1, argsLength), 1)
                Next
                data_r = data_r & "," & data_c
            Next
            .WriteLine Mid(data_r, 2)
        Next
    End With
    Debug.Print CDate(Now - d)
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'�w�肵���͈͓��̃����_���Ȑ������擾����.
'�ǂ����̃T�C�g���������Ă����R�[�h.
'-----------------------------------------
'����       :MinNum         :�����̍ŏ��l
'����       :MaxNum         :�����̍ő�l
'�߂�l     :�w��͈͓��̔C�ӂ̐���
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
Public Function RndScope(ByVal MinNum As Long, MaxNum As Long) As Long
    '�w�肵���͈̗͂����𐶐�
    Dim ret As Integer
    Randomize
    ret = Int(Rnd() * (MaxNum - MinNum + 1) + MinNum)
    RndScope = ret
End Function
