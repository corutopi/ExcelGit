Attribute VB_Name = "SS301_Database"
'dbÇÃéÌï Ç≤Ç∆Ç…ìKâûÇµÇΩSQLï∂Çî≠çsÇ≈Ç´ÇÈÇÊÇ§Ç…Ç∑ÇÈ


Public Function MakeSqlTalbeCreate(tableName As Variant, colNames As Variant, colTypes As Variant, Optional primaryKeys As Variant, Optional dbName As Variant)
    nameArray = makeSingleArray(colNames)
    typeArray = makeSingleArray(colTypes)
    
    Sql = "CREATE TABLE " & tableName
    ' columns
    datas = ""
    For i = LBound(nameArray) To UBound(nameArray)
        datas = datas & ", " & nameArray(i) & " " & typeArray(i)
    Next
    pk = ""
    ' primaryKeys
    If Not IsMissing(primaryKeys) Then
        For Each s In primaryKeys
            pk = pk & ", " & s
        Next
        pk = ", PRIMARY KEY(" & Mid(pk, 3) & ")"
    End If
    ' overall
    datas = Mid(datas, 3)
    Sql = Sql & " (" & datas & pk & ")"
    Sql = Sql & ";"
    
    MakeSqlTalbeCreate = Sql
End Function


Public Function MakeSqlInsertData(tableName As Variant, colTypes As Variant, colValues As Variant, Optional colNames As Variant)
    typeArray = makeSingleArray(colTypes)
    valueArray = makeSingleArray(colValues)
    
    Sql = "INSERT INTO " & tableName
    ' col names
    nms = ""
    If Not IsMissing(colNames) Then
        nameArray = makeSingleArray(colNames)
        For Each n In nameArray
            nms = nms & ", " & n
        Next
        nms = " (" & Mid(nms, 3) & ")"
    End If
    ' col values
    vls = ""
    For i = LBound(valueArray) To UBound(valueArray)
        v = valueArray(i)
        t = typeArray(i)
        If InStr(t, "VARCHAR") = 1 Or InStr(t, "TIMESTAMP") Then
            v = "'" & v & "'"
        End If
        vls = vls & ", " & v
    Next
    vls = " VALUES (" & Mid(vls, 3) & ")"
    ' overall
    Sql = Sql & nms & vls & ";"
    MakeSqlInsertData = Sql
End Function


Public Function MakeSqlDropTable(tableName As Variant)
    Sql = "DROP TABLE IF EXISTS " & tableName & ";"
    MakeSqlDropTable = Sql
End Function


'target is assumed for range
Private Function makeSingleArray(target As Variant) As Variant
    Dim s As Variant
    ReDim s(target.Count - 1)
    
    For i = 1 To target.Count
        s(i - 1) = target.Item(i).Value
    Next
    makeSingleArray = s
End Function


Public Sub testTables()
    'Debug.Print MakeTalbeCreateSQL(Range("B1"), Range("B3:D3"), Range("B4:D4"), Range("B3"))
End Sub

