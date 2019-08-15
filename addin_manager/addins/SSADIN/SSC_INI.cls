VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SSC_INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private z_path As String
Private z_sections As Dictionary

'Private Type Section
'    Name As String
'    Keys As Collection
'End Type
'Private Type Key
'    Name As String
'    Value As String
'End Type


Public Sub Read(path)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    '�f�[�^�`�F�b�N
    If Not fso.FileExists(path) Then
        'Err.Raise Number:=CP_ERR_NUMBER, Description:=CP_ERR_MSG_FileNotFound
    End If
    '�ϐ��ݒ�
    z_path = path
    Dim ts As Object
    Dim dataArr As Variant
    '�t�@�C���ǂݍ���
    Set ts = fso.OpenTextFile(path, 1)  'read
    dataArr = Split(ts.ReadAll, vbCrLf)
    ts.Close
    'INI�t�@�C���ǂݍ���
    Set z_sections = New Dictionary
    Dim keyDic As Dictionary
    Dim secName As String
    Dim s As String
    Dim i As Long
    For i = LBound(dataArr) To UBound(dataArr)
        s = Trim(dataArr(i))
        If Left(s, 1) = ";" Then
            '�R�����g�̔���
        ElseIf Left(s, 1) = "[" Then
            '�Z�N�V�����̓ǂݍ���
            If secName <> "" Then
                z_sections.Add secName, keyDic
            End If
            secName = Trim(Mid(s, 2, Len(s) - 2))
            Set keyDic = New Dictionary
        ElseIf Len(s) < Len(Replace(s, "=", "")) Then
            '�L�[�̓ǂݍ���
            '@todo ������ "=" ���܂ނƂ��̓�����l���ł��Ă��Ȃ�.
            Dim keyArr
            keyArr = Split(s, "=")
            keyDic.Add keyArr(0), keyArr(1)
        End If
    Next
    z_sections.Add secName, keyDic
End Sub


Public Function GetValue(Section, Key) As String
    If Not z_sections.Exists(Section) Then
        '@todo Raise error
    End If
    GetValue = z_sections.Item(Section).Item(Key)
End Function


Public Sub SetValue(Section, Key, Value)
    Dim d As Dictionary
    If Not z_sections.Exists(Section) Then
        Set d = New Dictionary
        z_sections.Add Section, d
    End If
    If z_sections.Item(Section).Exists(Key) Then
        z_sections.Item(Section).Remove Key
    End If
    z_sections.Item(Section).Add Key, Value
End Sub


Public Sub RemoveValue(Section, Key)
End Sub


Public Sub DropSection(Section)
End Sub


Public Sub Export(path)
End Sub


Public Sub PrintData()
End Sub