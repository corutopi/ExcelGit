Attribute VB_Name = "SS201_VBComponent"
'Microsoft Visual Basic for Applications Extensibility�̎Q�Ɛݒ肪�K�v

Option Explicit

'*******************************************************************************
'--exportAllModule-------------------------
'�I�������t�@�C�����̃��W���[�������ׂăG�N�X�|�[�g����B
'-----------------------------------------
'����       :�Ώۃt�@�C���p�X
'����       :�o�͐�t�H���_�B�����l�͑Ώۃt�@�C���Ɠ����t�H���_
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20190805   :xxx            :�������t�@�C���p�X����WorkbookObj�ɕύX
'20180701   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub ExportAllModule(wb As Workbook, Optional exportFolder = "")
    Dim tm              As Object           'targetModule
    Dim fileExt         As String
    
    If exportFolder = "" Then exportFolder = wb.path
    
    For Each tm In wb.VBProject.VBComponents
        fileExt = GetExtFromModuleType(tm.Type)
        If fileExt <> "" Then
            Call ExportModule(tm, exportFolder, fileExt)
        End If
    Next
    
    Exit Sub
End Sub

'*******************************************************************************
'--getExtFromModuleType-------------------------
'���W���[���^�C�v(VBComponent.Type)���Ƀe�L�X�g�t�@�C���o�͎���
'�g�p����g���q���擾����
'-----------------------------------------
'����       :mType          :�Ώۓ�
'�߂�l     :�Ή�����g���q
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171220   :xxx            :�V�K�쐬
'*******************************************************************************
Public Function GetExtFromModuleType(ByVal mType As Integer) As String
    Dim re As String
    
    re = ""
    Select Case mType
        Case 1  'vbext_ct_StdModule
            re = "bas"
        Case 2, 100 'vbext_ct_ClassModule, vbext_ct_Document
            '�h�L�������g���W���[���̓G�N�X�|�[�g�Ώۂ���O���Ă���������
            re = "cls"
        Case 3  'vbext_ct_MSForm
            re = "frm"
    End Select

    GetExtFromModuleType = re
End Function

'*******************************************************************************
'--exportModule-------------------------
'���W���[�����o�͂���
'-----------------------------------------
'����       :target         :�o�̓��W���[���R���|�[�l���g
'����       :path           :�o�͐�p�X
'����       :ext            :�o�͊g���q
'�߂�l     :�Ή�����g���q
'-----------------------------------------
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20171220   :xxx            :�V�K�쐬
'*******************************************************************************
Public Sub ExportModule(ByVal target As Object, ByVal path As String, ByVal ext As String)
    target.Export path & "\" & target.Name & "." & ext
End Sub


'*******************************************************************************
'--comment_version_0.1.0------------------
'���W���[�����C���|�[�g����.
'���ɃC���|�[�g����Ă��郂�W���[���̏ꍇ�͏㏑�����邩�ʖ��ŃC���|�[�g����.
'-----------------------------------------
'����       :hikihiki       :�����̐���
'�߂�l     :�߂�l�̓��e
'-----------------------------------------
'@todo �o�b�N�A�b�v���邩�ۂ���������֗�����
'--�X�V����-------------------------------
'yyyymmdd   :xxx            :[�X�V���e]
'20190804   :               :�V�K�쐬
'*******************************************************************************
Public Sub ImportModule(modulePath As String, targetBook As Workbook, Optional overWrite As Boolean = True)
    '�����̃��W���[�������m�F
    If overWrite Then
        Dim moduleName As String
        Dim moduleExt As String
        Dim tm As Object  'VBComponent
        moduleName = Mid(modulePath, InStrRev(modulePath, "\") + 1)
        moduleName = Mid(moduleName, 1, InStrRev(moduleName, ".") - 1)  '�g���q�̏���
        moduleExt = Mid(modulePath, InStrRev(modulePath, ".") + 1)
        For Each tm In targetBook.VBProject.VBComponents
            '���ꃂ�W���[������
            If tm.Name <> moduleName Then
                '���̂���v���Ă��邱�� -> �t�@�C�������ύX����Ă��Ȃ����Ƃ��O��
            ElseIf GetExtFromModuleType(tm.Type) <> moduleExt Then
                '�g���q����v���邱��
            ElseIf tm.Type = 100 Then
                '�h�L�������g�łȂ�����
            Else
                targetBook.VBProject.VBComponents.Remove tm
                Exit For
            End If
        Next
    End If
    '�C���|�[�g
    targetBook.VBProject.VBComponents.Import modulePath
End Sub


