Attribute VB_Name = "AD102_ForWorksheet"
Option Explicit


'*******************************************************************************
'-----------------------------------------
'�A�N�e�B�u��Ԃ̃u�b�N�̂��ׂĂ̕\�����V�[�g�̑I���Z������ѕ\��
'��Ԃ�A1�Z���ɐ��ڂ���B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171230   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�u�b�N����()
    Dim ws As Worksheet
    Dim ti As Integer           'tmp index
    
    On Error GoTo ERR_MANAGER
    
    Call startProcess
    
    ti = 0
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If ti = 0 Then
                ti = ws.Index
            End If
            ws.Select
            Application.GoTo Range("A1"), True
        End If
    Next ws
    
    ActiveWorkbook.Worksheets(ti).Select
    
    Call endProcess
    Exit Sub
    
ERR_MANAGER:
    Call errProcess
End Sub


'*******************************************************************************
'-----------------------------------------
'�A�N�e�B�u��Ԃ̃u�b�N�̂��ׂĂ̔�\���V�[�g��\����Ԃɂ���B
'�u�b�N�ی�ɂ�鏈�����s�ɂ͖��Ή��B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171230   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_��\���V�[�g�S�\��()
    Dim ws As Worksheet
    Dim ti As Integer           'tmp index
    
    On Error GoTo ERR_MANAGER
    
    Call startProcess
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
    Next ws
    
    Call endProcess
    Exit Sub
    
ERR_MANAGER:
    Call errProcess
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃u�b�N�̃V�[�g����I�𒆂̃Z�����牺�����ɗ񋓂���B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180119   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�V�[�g�����o()
    Dim tb As Workbook
    Dim ts As Worksheet
    Dim ws As Worksheet
    Dim tr As Range
    
    Dim msg As String
    Dim r As Long
    Dim c As Long
    
    msg = "�I�𒆂̃Z�����牺�����ɃA�N�e�B�u�u�b�N���̃V�[�g����񋓂��܂��B" _
            & vbNewLine & "���s���Ă�낵���ł����H"
            
    If MsgBox(msg, vbYesNo) = vbYes Then
        Set tb = ActiveWorkbook
        Set ts = ActiveSheet
        Set tr = Selection
        r = tr.Item(1).row
        c = tr.Item(1).column
        
        For Each ws In tb.Worksheets
            ts.Cells(r, c).Value = ws.Name
            r = r + 1
        Next
    End If
End Sub


'*******************************************************************************
'-----------------------------------------
'�I�𒆂̃u�b�N�̃V�[�g����ύX����B
'�I�𒆂̃Z���͈͂�1��ڂ̃V�[�g����2��ڂ̖��O�ɕύX����B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180119   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�V�[�g���ύX()
    Dim tb As Workbook
    Dim ts As Worksheet
    Dim ws As Worksheet
    Dim tr As Range
    
    Dim msg As String
    Dim r As Long
    Dim c As Long
    
    msg = "�I�𒆂̃Z���������ɃV�[�g����ύX���܂��B" _
            & vbNewLine & "1��ڂ̃V�[�g����2��ڂ̖��O�ɕύX���܂��B" _
            & vbNewLine & "" _
            & vbNewLine & "��낵���ł����H"
            
    If MsgBox(msg, vbYesNo) = vbYes Then
        Set tb = ActiveWorkbook
        Set ts = ActiveSheet
        Set tr = Selection
        
        For r = 1 To tr.Rows.Count
            For Each ws In ActiveWorkbook.Worksheets
                If tr.Item(r, 1).Value = ws.Name And tr.Item(r, 2).Value <> "" Then
                    ws.Name = tr.Item(r, 2).Value
                    Exit For
                End If
            Next
        Next
    End If
End Sub


'*******************************************************************************
'--SA_�V�[�g���я��ύX-------------------------
'�I�𒆂̃u�b�N�̃V�[�g����ύX����B
'�I�𒆂̃Z���͈͂�1��ڂ̃V�[�g����2��ڂ̖��O�ɕύX����B
'-----------------------------------------
'����       :�Ȃ�
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20180205   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub SA_�V�[�g���я��ύX()
    Dim tb As Workbook
    Dim ts As Worksheet
    Dim ws As Worksheet
    Dim tr As Range
    
    Dim msg As String
    Dim r As Long
    Dim c As Long
    
    msg = "�I�𒆂̃Z���������ɃV�[�g�̕��я���ύX���܂��B" _
            & vbNewLine & "1��ڂɋL�ڂ̃V�[�g���ォ�珇�ɐ擪�ɂȂ�悤�ɕ��ׂ܂��B" _
            & vbNewLine & "" _
            & vbNewLine & "��낵���ł����H"
            
    If MsgBox(msg, vbYesNo) = vbYes Then
        Set tb = ActiveWorkbook
        Set ts = ActiveSheet
        Set tr = Selection
        
        For r = tr.Rows.Count To 1 Step -1
            For Each ws In ActiveWorkbook.Worksheets
                If tr.Item(r, 1).Value = ws.Name Then
                    ws.Move tb.Worksheets(1)
                    Exit For
                End If
            Next
        Next
    End If
End Sub



