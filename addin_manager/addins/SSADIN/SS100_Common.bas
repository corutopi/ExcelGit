Attribute VB_Name = "SS100_Common"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub startProcess(Optional ByVal Name As String = "")
    Application.ScreenUpdating = False
    Call logging("�����J�n�F" & Name)
End Sub

Public Sub endProcess(Optional ByVal Name As String = "")
    Call logging("�����I���F" & Name)
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub logging(ByVal msg As String)
    Application.StatusBar = msg
End Sub

Public Sub errProcess(Optional ByVal Description As String = "", Optional ByVal isEnd As Boolean = False)
    Dim msg As String
    
    If Description <> "" Then
        '�G���[�������ʂɐݒ肳��Ă���ꍇ
        msg = Description
    ElseIf Err.Number <> 0 Then
        '�V�X�e���G���[�ŌĂяo���ꂽ�ꍇ
        msg = "�G���[���������܂����B" _
                & vbNewLine & "�G���[�ԍ��F" & Err.Number _
                & vbNewLine & "�G���[���e�F" & Err.Description
    Else
        '�G���[�������Ȃ��V�X�e���G���[�ł��Ȃ��ꍇ
        msg = "�G���[���肪�s���܂������A�ڍׂ����ʂł��܂���ł����B"
    End If
    
    MsgBox msg, vbExclamation, "���ʃG���[����"
    
    If isEnd Then
        End     '�����I��
    End If
End Sub

Public Sub errTest()
    Debug.Print "start"
    
    Call errProcess
    
    Debug.Print "end"
End Sub
