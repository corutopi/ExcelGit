Attribute VB_Name = "★試し"
Public Sub aaaa()
    Dim df As New SSC_MyDataFrame
    df.ReadCSV "D:\Corutopi\040_Program\ExcelVBA\カレンダーデータ.csv"
    df.ReadCell Selection
    Stop
End Sub





Public Sub onerrortest()
    'On Error Resume Next
    
    Call onerrortestsub
End Sub


Public Sub onerrortestsub()
    'On Error GoTo X
    Err.Raise 8
    Exit Sub
X:
    Debug.Print "onerrortestsub error"
End Sub





Public Sub testt()
    path = "D:\Corutopi\040_Program\ExcelVBA\csv_test.csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    stt = fso.OpenTextFile(path, 1).ReadAll()
    Debug.Print Len(stt)
    d = Now
    Set r = GetCSVCollection(CStr(stt))
    Debug.Print CDate(Now - d)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Cells.Clear
    For i = 1 To r.Count
        Set c = r.Item(i)
        For j = 1 To c.Count
            ws.Cells(i, j).Value = c.Item(j)
        Next
    Next
End Sub



Public Function hogehoge(csvData As String)
    Debug.Print Len(csvData)
    d = Now()
    For i = 1 To Len(csvData)
        j = Mid(csvData, i, 1)
        Debug.Print j
    Next i
    Debug.Print CDate(Now - d)
    
End Function



