Attribute VB_Name = "SS701_特殊_海老蔵君"
Option Explicit
 
Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClipboardViewer Lib "user32.dll" (ByVal hWndNewViewer As LongPtr) As LongPtr
Private Declare PtrSafe Function ChangeClipboardChain Lib "user32.dll" (ByVal hWndRemove As LongPtr, ByVal hWndNewNext As LongPtr) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal format As Long) As Long
 
Private Const GWL_WNDPROC As Long = -4
 
Private Const WM_DRAWCLIPBOARD As Long = &H308
Private Const WM_CHANGECBCHAIN As Long = &H30D
Private Const WM_NCHITTEST As Long = &H84
 
Private Const CF_BITMAP As Long = 2
 
Private Const ROW_HEIGHT As Double = 13.5
 
Private hWndForm As LongPtr
Private wpWindowProcOrg As Long
Private hWndNextViewer As LongPtr
Private firstFired As Boolean

'Private Const eviSheetName As String = "エビデンス"

Private eviSheet As Worksheet

Public Sub catchClipboard(ByVal ws As Worksheet)
    Set eviSheet = ws

    'ハンドラーの設定
    hWndForm = FindWindow("ThunderDFrame", SS701_Evidence.Caption)
    'ウィンドウプロシージャの設定
    wpWindowProcOrg = SetWindowLong(hWndForm, GWL_WNDPROC, AddressOf WindowProc)
    '監視開始直後の誤動作を防ぐ
    firstFired = False
    'クリップボードのイベントを受け取るクリップボードビューアの設定
    hWndNextViewer = SetClipboardViewer(hWndForm)
End Sub
 
Public Sub releaseClipboard()
    'フォームをクリップボードから切り離し
    Call ChangeClipboardChain(hWndForm, hWndNextViewer)
    'ウィンドウプロシージャの設定をもとに戻す
    Call SetWindowLong(hWndForm, GWL_WNDPROC, wpWindowProcOrg)
    
    Set eviSheet = Nothing
End Sub
 
Public Function WindowProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Select Case uMsg
        Case WM_DRAWCLIPBOARD
            If Not firstFired Then
                firstFired = True
            ElseIf IsClipboardFormatAvailable(CF_BITMAP) <> 0 Then
                pasteToSheet
            End If
            If hWndNextViewer <> 0 Then
                Call SendMessage(hWndNextViewer, uMsg, wParam, lParam)
            End If
            WindowProc = 0
        Case WM_CHANGECBCHAIN
            If wParam = hWndNextViewer Then
                hWndNextViewer = lParam
            ElseIf hWndNextViewer <> 0 Then
                Call SendMessage(hWndNextViewer, uMsg, wParam, lParam)
            End If
            WindowProc = 0
        Case WM_NCHITTEST
            WindowProc = 0
        Case Else
            WindowProc = CallWindowProc(wpWindowProcOrg, hWndForm, uMsg, wParam, lParam)
    End Select
End Function
 
Private Sub pasteToSheet()
    Dim rowIdx As Integer
    Dim tempSheet As Worksheet
    Dim check As Boolean
    
    check = False
    
'    'エビデンス用シートの確認。無ければ作成する。
'    For Each tempSheet In ThisWorkbook.Worksheets
'        If tempSheet.name = eviSheetName Then
'            check = True
'            Exit For
'        End If
'    Next
'
'    If check = False Then
'        ThisWorkbook.Worksheets.Add
'        ActiveSheet.name = eviSheetName
'    End If
    
    If Application.ClipboardFormats(1) = xlClipboardFormatBitmap Then
        'コピー対象がビットマップ画像のときのみ貼り付ける
        With eviSheet
            If .Shapes.Count > 0 Then
                With .Shapes(.Shapes.Count)
                    rowIdx = (.Top + .Height) / ROW_HEIGHT + 4
                End With
            Else
                rowIdx = 1
            End If
            .Cells(rowIdx, 1).PasteSpecial
        End With
    End If
End Sub

