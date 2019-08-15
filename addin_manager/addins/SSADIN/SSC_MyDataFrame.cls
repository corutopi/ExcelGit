VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SSC_MyDataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'--comment_version_0.1.0------------------
'2�����f�[�^�������₷�����邽�߂̃N���X.
'�s��f�[�^��2�d��Collection�N���X�ŕێ���,�e�֐����g�p���ăA�N�Z�X����B
'
'
'
'-----------------------------------------
'����       :---            :
'�߂�l     :---
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************

Option Explicit

'�G���[�֘A
Private Const CP_ERR_NUMBER As Long = 9999
Private Const CP_ERR_MSG_FileNotFound As String = "�t�@�C�������݂��܂���."
Private Const CP_ERR_MSG_CellAreaNum As String = "�����̃Z���͈͎͂w��ł��܂���."
Private Const CP_ERR_MSG_ArrayData As String = "�����Ɏw��ł���̂�Collection���I�u�W�F�N�g��v�f�Ɋ܂܂Ȃ�Variant�z��݂̂ł�."
Private Const CP_ERR_MSG_DataColumnNum As String = "�f�[�^�񐔂���v���܂���."
Private Const CP_ERR_MSG_OptionArgs As String = "���� {%s} �͉��ꂩ {%n}�����w��ł��܂���."
'�N���X���񋓌^
Private Enum EP_VALUE
    mdfNone
    mdfValueCollection
    mdfValueArray
    mdfValue
    mdfObject
End Enum

'�v�f
Private z_header As New Collection
Private z_datas As New Collection
Private z_index As New Collection  '������


'*******************************************************************************
'--comment_version_0.1.0------------------
'CSV�t�@�C����ǂݍ���
'
'-----------------------------------------
'����       :path           :CSV�t�@�C���p�X
'����       :hasHeader      :�w�b�_�[�̗L��
'�߂�l     :---
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
'CSV�t�@�C����ǂݍ���
Public Sub ReadCSV(path As String, Optional hasHeader As Boolean = True)
    Dim fso As Object
    Dim ts As Object
    Dim dataArr As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    '�f�[�^�`�F�b�N
    If Not fso.FileExists(path) Then
        Err.Raise Number:=CP_ERR_NUMBER, Description:=CP_ERR_MSG_FileNotFound
    End If
    '������
    Call Reset
    '�t�@�C���I�[�v��
    Set ts = fso.OpenTextFile(path, 1)  'read
    dataArr = Split(ts.ReadAll, vbCrLf)
    ts.Close
    '�񐔌���
    Dim i As Long
    Dim maxCol As Long
    Dim datas As Variant
    For i = LBound(dataArr) To UBound(dataArr)
        datas = Split(dataArr(i), ",")
        If maxCol <= UBound(datas) Then
            maxCol = UBound(datas) + 1
        End If
    Next
    '�f�[�^�쐬
    Dim tmpData As Collection
    Dim j As Long
    For i = LBound(dataArr) To UBound(dataArr)
        datas = Split(dataArr(i), ",")
        If i = 1 And hasHeader Then
            For j = LBound(datas) To UBound(datas)
                z_header.Add datas(j)
            Next
            For j = UBound(datas) + 1 To maxCol - 1
                z_header.Add ""
            Next
        Else
            Set tmpData = New Collection
            For j = LBound(datas) To UBound(datas)
                tmpData.Add datas(j)
            Next
            For j = UBound(datas) + 1 To maxCol - 1
                tmpData.Add ""
            Next
            z_datas.Add tmpData
        End If
    Next
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'�Z���f�[�^��ǂݍ���
'
'-----------------------------------------
'����       :target         :�Ǎ��Z���͈�
'����       :hasHeader      :�w�b�_�[�̗L��
'�߂�l     :---
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
Public Sub ReadCell(target As Range, Optional hasHeader As Boolean = True)
    Dim r As Long
    Dim c As Long
    Dim tmpData As Collection
    
    '�f�[�^�`�F�b�N
    If target.Areas.Count > 1 Then
        Err.Raise Number:=CP_ERR_NUMBER, Description:=CP_ERR_MSG_CellAreaNum
    End If
    '������
    Call Reset
    '�f�[�^�쐬
    For r = 1 To target.Rows.Count
        If r = 1 And hasHeader Then
            For c = 1 To target.Columns.Count
                z_header.Add target.Item(r, c).Value
            Next
        Else
            Set tmpData = New Collection
            For c = 1 To target.Columns.Count
                tmpData.Add target.Item(r, c).Value
            Next
            z_datas.Add tmpData
        End If
    Next
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'CSV�t�@�C���ɏo�͂���
'
'-----------------------------------------
'����       :path           :�o�͐�t�@�C���p�X
'�߂�l     :---
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
Public Sub ExportCSV(path)
    Dim fso As Object
    Dim tmpStr As String
    Dim i As Long
    Dim j As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.CreateTextFile(path)
        '�w�b�_�[�̏o��
        tmpStr = ""
        If z_header.Count > 0 Then
            For i = 1 To Me.ColumnsCount
                tmpStr = tmpStr & "," & z_header.Item(i)
            Next
        End If
        .WriteLine Mid(tmpStr, 2)
        '�f�[�^�̏o��
        For i = 1 To Me.Count
            tmpStr = ""
            For j = 1 To Me.ColumnsCount
                tmpStr = tmpStr & "," & z_datas.Item(i).Item(j)
            Next
            .WriteLine Mid(tmpStr, 2)
        Next
    End With
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'�f�[�^�����Z�b�g����.
'�ێ����Ă���w�b�_�[���, �f�[�^�������ׂĔj������.
'-----------------------------------------
'����       :---            :
'�߂�l     :---
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
Public Sub Reset()
    Set z_header = New Collection
    Set z_datas = New Collection
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'�f�[�^�̍s�����擾����.
'
'-----------------------------------------
'����       :---            :
'�߂�l     :�s��
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
Public Property Get Count()
    Count = z_datas.Count
End Property


'*******************************************************************************
'--comment_version_0.1.0------------------
'�f�[�^�̗񐔂��擾����.
'
'-----------------------------------------
'����       :---            :
'�߂�l     :��
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
Public Function ColumnsCount()
    Dim re As Long
    
    If z_header.Count <> 0 Then
        re = z_header.Count
    ElseIf z_datas.Count <> 0 Then
        re = z_datas.Item(1).Count
    Else
        re = 0
    End If
    ColumnsCount = re
End Function


'*******************************************************************************
'--comment_version_0.1.0------------------
'�s�f�[�^��ǉ�����
'
'-----------------------------------------
'����       :---            :
'�߂�l     :��
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'*******************************************************************************
Public Sub AddRow(Data, Optional Before, Optional After)
    '@todo ���N���X�^�̃f�[�^���A�g���ꂽ�ꍇ�̓���
    Dim tmpColl As New Collection
    Dim tmpColl2 As New Collection
    Dim i As Long
    
    '�f�[�^�̃`�F�b�N
    Set tmpColl = ToCollection(Data)
    If tmpColl Is Nothing Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=CP_ERR_MSG_ArrayData
    End If
    If tmpColl.Count <> Me.Count Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=CP_ERR_MSG_DataColumnNum
    End If
    If IsMissing(Before) And IsMissing(After) Then
    ElseIf Not IsMissing(Before) And Not IsMissing(After) Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=CP_ERR_MSG_OptionArgs
    End If
    '�f�[�^�ǉ�(�z��/�R���N�V�����̏ꍇ)
    For i = 1 To tmpColl.Count
        tmpColl2.Add tmpColl.Item(i), , Before, After
    Next
    z_datas.Add tmpColl
End Sub


Public Sub AddCol()

End Sub


'�w�b�_�[�̐ݒ�
Public Sub SetHeader(header)
    Dim tmpColl As Collection
    
    '�G���[�`�F�b�N
    Set tmpColl = ToCollection(header)
    If tmpColl Is Nothing Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=""
    End If
    If Not (z_header.Count = 0 And z_datas.Count = 0) And _
            ColumnsCount <> tmpColl.Count Then
        Err.Raise Number:=CP_ERR_NUMBER, _
                  Description:=CP_ERR_MSG_DataColumnNum
    End If
    '�w�b�_�[�̍Đݒ�
    Set z_header = tmpColl  'ToCollection�ŐV�����I�u�W�F�N�g�ɂȂ��Ă邩��OK
End Sub


'�f�[�^�̒��o
Public Function GetValue(row, column)
    Dim tmpRow As Long
    Dim tmpColumn As Long
    
    tmpRow = row
    If column = Int(column) Then
        tmpColumn = column
    Else
        tmpColumn = GetColumnNum(column)
    End If
    GetValue = z_datas.Itemm(tmpRow).Item(tmpColumn)
End Function


'�f�[�^�̍X�V
Public Sub SetValue(row, column, val)
    Dim tmpRow As Long
    Dim tmpColumn As Long
    
    tmpRow = row
    If column = Int(column) Then
        tmpColumn = column
    Else
        tmpColumn = GetColumnNum(column)
    End If
    z_datas.Itemm(tmpRow).Add val, Before:=column
    z_datas.Itemm(tmpRow).Remove column + 1
End Sub


'�f�[�^�s�̍폜
Public Sub Delete(row)
    Dim tmpRow As Long
    
    tmpRow = row
    z_datas.Remove tmpRow
End Sub


'�w�b�_�[�̖��́̔ԍ��ϊ�
Public Function GetColumnNum(Name) As Long
    Dim re As Long
    Dim i As Long
    
    re = -1
    For i = 1 To z_header.Count
        If z_header.Item(i) = Name Then
            re = i
            Exit For
        End If
    Next
    GetColumnNum = re
End Function


'�f�[�^�̘A���i�s�j


'�f�[�^�̃\�[�g
Public Function Sort(Columns, Optional IsAsc As Boolean = True) As SSC_MyDataFrame
    Dim re As New SSC_MyDataFrame
    Dim tmpColl As New Collection
    Dim sortTarget As Variant
    Dim i As Long
    Dim j As Long
    
    Set tmpColl = ToCollection(Columns)
    sortTarget = tmpColl.Item(tmpColl.Count) '�Ō�̃f�[�^�Ń\�[�g����
    For i = 1 To Me.Count
        Dim beforePoint As Long
        beforePoint = -1
        For j = 1 To re.Count
            If (Me.GetValue(i, sortTarget) < re.GetValue(j, sortTarget)) = IsAsc Then
                beforePoint = j
                Exit For
            End If
        Next
        If beforePoint = -1 Then
            '1�s�ۂ��Ƃ̃f�[�^���ǂ��擾���邩�H
            re.AddRow z_datas.Item(j)
        Else
            re.AddRow z_datas.Item(j), Before:=beforePoint
        End If
    Next
    
    If tmpColl.Count > 1 Then
        tmpColl.Remove tmpColl.Count
        Set re = re.Sort(tmpColl, IsAsc)
    End If
    Set Sort = re
End Function


'�f�[�^�̈�Ӊ�
Public Function UniqueData(Optional Columns) As SSC_MyDataFrame
    Dim tmpColl As New Collection
    Dim columnNums As New Collection
    Dim re As SSC_MyDataFrame
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    '�m�F�Ώۗ�̌���
    If IsMissing(Columns) Then
        For i = 1 To Me.Count
            columnNums.Add i
        Next
    Else
        Set tmpColl = ToCollection(Columns)
        For i = 1 To tmpColl.Count
            columnNums.Add Me.GetColumnNum(tmpColl.Item(i))
        Next
    End If
    
    '�d���f�[�^�̍폜
    Set re = Me.Copy()
    i = 1
    Do While i <= re.Count
        Dim check As Boolean
        check = True  '�m�F�Ώۗ񂪂��ׂĈ�v����f�[�^�������True
        For j = i - 1 To 1 Step -1
            For k = 1 To columnNums.Count
                If re.GetValue(i, k) <> re.GetValue(j, k) Then
                    check = False
                    Exit For
                End If
            Next
            If check = False Then Exit For
        Next
        If check Then
            re.Delete i
        Else
            i = i + 1
        End If
    Loop
    Set UniqueData = re
End Function


'�f�[�^�̑I��
Public Function SelectData(column, Value, Optional isEqual = True) As SSC_MyDataFrame
    Dim columnNum As Long
    Dim re As SSC_MyDataFrame
    Dim i As Long
    
    Set re = Me.Copy()
    i = 1
    Do While i <= re.Count
        If re.GetValue(i, column) <> Value Then
            re.Delete i
        End If
    Loop
    Set SelectData = re
End Function


'�f�[�^���R�s�[����
Public Function Copy() As SSC_MyDataFrame
    Dim re As New SSC_MyDataFrame
    Dim i As Long
    
    If z_header.Count > 0 Then
        re.SetHeader ToCollection(z_header)
    End If
    For i = 1 To Me.Count
        re.AddRow ToCollection(z_datas.Item(i))
    Next
    
    Set Copy = re
End Function


'�f�[�^���o�͂���
Public Sub PrintData()
    Dim maxChara As New Collection  '�s���Ƃ̍ő�o�͕�����
    Dim r As Long
    Dim c As Long
    Dim tmpChara As Long
    Dim tmpStr As String
    
    '�ő啶��������
    For c = 1 To Me.ColumnsCount
        tmpChara = 0
        If z_header.Count > 0 Then
            tmpChara = GetByteCharaNum(z_header.Item(c))
        End If
        For r = 1 To z_datas.Count
            If tmpChara < GetByteCharaNum(z_datas.Item(r).Item(c)) Then
                tmpChara = GetByteCharaNum(z_datas.Item(r).Item(c))
            End If
        Next
        maxChara.Add tmpChara
    Next
    '�f�[�^�o��
    tmpStr = ""
    If z_header.Count > 0 Then
        For c = 1 To z_header.Count
            tmpStr = tmpStr & "," & String(maxChara.Item(c) - GetByteCharaNum(z_header.Item(c)), " ") & z_header.Item(c)
        Next
        Debug.Print Mid(tmpStr, 2)
    Else
        Debug.Print "[None Header]"
    End If
    For r = 1 To z_datas.Count
        tmpStr = ""
        For c = 1 To z_datas.Item(r).Count
            tmpStr = tmpStr & "," & String(maxChara.Item(c) - GetByteCharaNum(z_datas.Item(r).Item(c)), " ") & z_datas.Item(r).Item(c)
        Next
        Debug.Print Mid(tmpStr, 2)
    Next
    If z_datas.Count = 0 Then
        Debug.Print "[None Data]"
    End If
End Sub


'*******************************************************************************
'��������N���X�����\�b�h*******************************************************
'*******************************************************************************
'�o�C�g���������擾����
Private Function GetByteCharaNum(str) As Long
    GetByteCharaNum = LenB(StrConv(str, vbFromUnicode))
End Function


'�z��/�P��̒l���R���N�V�����^�ɕϊ�����
Private Function ToCollection(val) As Collection
    Dim re As New Collection
    Dim valType As EP_VALUE
    Dim i As Long
    
    valType = GetValueType(val)
    
    If valType = mdfValue Then
        re.Add val
    ElseIf valType = mdfValueArray Then
        For i = LBound(val) To UBound(val)
            re.Add val(i)
        Next
    ElseIf valType = mdfValueCollection Then
        For i = 1 To val.Count
            re.Add val.Item(i)
        Next
    Else
        Set re = Nothing
    End If
    Set ToCollection = re
End Function


'���������̃N���X�Ŏg�p�\�Ȕz��/�R���N�V����/�P��̒l�ł��邱�Ƃ��m�F����
Private Function GetValueType(obj) As EP_VALUE
    Dim re As EP_VALUE
    Dim arrCheck As EP_VALUE
    Dim i As Long
    re = mdfNone
    If TypeName(obj) = "Collection" Then
        '�R���N�V����
        arrCheck = mdfValueCollection
        For i = 1 To obj.Count
            If IsObject(obj.Item(i)) Then
                arrCheck = mdfObject
                Exit For
            End If
        Next
        re = arrCheck
    ElseIf IsArray(obj) Then
        '�z��
        arrCheck = mdfValueArray
        For i = LBound(obj) To UBound(obj)
            If IsObject(obj(i)) Then
                arrCheck = mdfObject
                Exit For
            End If
        Next
        re = arrCheck
    ElseIf IsObject(obj) Then
        '�P��̒l
        re = mdfValue
    Else
        '���̂ق��̒l�͕s��
        re = mdfObject
    End If
    
    GetValueType = re
End Function


'�G���[���b�Z�[�W���쐬���� @todo ������
Private Function MakeErrMsg(msg, s, n) As String

End Function
