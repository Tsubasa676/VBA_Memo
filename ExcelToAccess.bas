Option Explicit
' ####################################
' ###  Accessに関する処理を行う    ###
' ####################################
 
' SQL結果をカラム配列、データ配列に格納
Function SelectTable(ByVal accdbPath As String, ByVal sql As String, ByRef arrCol As Variant, ByRef arrData As Variant) As Boolean
    SelectTable = False
    Application.StatusBar = "データ取得実行：" & sql
    ' エラー処理
    On Error GoTo EXE_ERR
    ' データベース設定
    Dim accCn As New ADODB.Connection
    Dim accRs As New ADODB.Recordset
    accCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accdbPath & ";"
    accRs.CursorLocation = 3
    accRs.Open sql, accCn
    ' カウント
    Dim i As Long, j As Long
    ' SQL結果格納
    If accRs.RecordCount > 0 Then
        Dim temp As Variant
        Dim tempCol As Variant
        Dim tempRs As Variant
        temp = accRs.GetRows
        ReDim tempCol(1 To 1, 1 To UBound(temp) + 1)
        ReDim tempRs(1 To UBound(temp, 2) + 1, 1 To UBound(temp) + 1)
        For i = 0 To UBound(temp, 2)
            For j = 0 To UBound(temp)
                tempCol(1, j + 1) = accRs.Fields.Item(j).Name
                tempRs(i + 1, j + 1) = temp(j, i)
            Next
        Next
        arrCol = tempCol
        arrData = tempRs
    Else
        Dim tempFields As ADODB.Fields
        Set tempFields = accRs.Fields
        ReDim tempCol(1 To 1, 1 To tempFields.Count)
        For i = 0 To (tempFields.Count - 1)
            tempCol(1, i + 1) = tempFields.Item(i).Name
        Next
        arrCol = tempCol
    End If
    ' DB切断
    accCn.Close
    Set accCn = Nothing
    SelectTable = True
    Exit Function
EXE_ERR:
    Call Common.ErrorMessage("SelectTable")
End Function
 
' テーブルにInsert文を実行する
' 例：Call accDB.InsertTable(path, "INSERT INTO TableName (文字列, 数値) VALUES ('test2', 123)")
Function InsertTable(ByVal accdbPath As String, ByVal sql As String) As Boolean
    InsertTable = False
    ' エラー処理
    On Error GoTo EXE_ERR
    ' データベース設定
    Dim accCn As New ADODB.Connection
    Dim accRs As New ADODB.Recordset
    accCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accdbPath & ";"
    accRs.CursorLocation = 3
    accCn.Execute sql
    ' DB切断
    accCn.Close
    Set accCn = Nothing
    InsertTable = True
    Exit Function
EXE_ERR:
    Call Common.ErrorMessage("InsertTable")
End Function
 
' テーブルにExcelファイルのデータを追加する
Function InsertTableFormExcel(ByVal accdbPath As String, ByVal tableName As String, _
                            ByVal xlsxPath As String, ByVal startCell As String, ByVal sheetNo As Integer) As Boolean
    InsertTableFormExcel = False
    Application.StatusBar = "「" & tableName & "」テーブルにデータ追加実行..."
    ' エラー処理
    'On Error GoTo EXE_ERR
    Dim errInfo As String
    ' Excelファイル読み込み
    Dim wb As Workbook
    Dim xData As Variant
    Set wb = Workbooks.Open(fileName:=xlsxPath, ReadOnly:=True)
    wb.Sheets(sheetNo).Activate
    xData = wb.Worksheets(sheetNo).Range(startCell, ActiveCell.SpecialCells(xlLastCell))
    wb.Close
    Application.StatusBar = "「" & tableName & "」テーブルに" & UBound(xData) & " 件のデータ追加実行..."
 
    ' データベース設定
    Dim accCn As New ADODB.Connection
    Dim accRs As New ADODB.Recordset
    Dim accRs2 As New ADODB.Recordset
    Dim accCol As String
    Dim accColumn As Variant
    Dim i As Long, j As Long
    accCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accdbPath & ";"
    accRs.CursorLocation = 3
    ' データベースのカラム取得（ID列を除外する）
    accRs.Open tableName, accCn
    For i = 0 To (accRs.Fields.Count - 1)
        If i > 1 Then
            accCol = accCol & "," & accRs.Fields(i).Name
        Else
            accCol = accRs.Fields(i).Name
        End If
    Next
    accColumn = Split(accCol, ",")
    accRs.Close
    ' Recordsetテンプレ作成
    With accRs2.Fields
        For i = 0 To UBound(accColumn)
            .Append accColumn(i), adVarWChar, -1
        Next
    End With
    ' データ追加
    accRs2.Open tableName, accCn, adOpenKeyset, adLockOptimistic
    For i = 1 To UBound(xData)
        accRs2.AddNew
        For j = 0 To UBound(accColumn)
            errInfo = "テーブル：" & tableName & vbCrLf & "行：" & i & vbCrLf & "列：" & accColumn(j) & vbCrLf & "データ：" & xData(i, j + 1)
            If Not IsEmpty(xData(i, j + 1)) Then
                accRs2.Fields(accColumn(j)) = xData(i, j + 1)
            End If
        Next
        accRs2.Update
    Next
    accRs2.Close
    ' DB切断
    accCn.Close
    Set accCn = Nothing
    InsertTableFormExcel = True
    Exit Function
EXE_ERR:
    Call Common.ErrorMessage("ImportTable", errInfo)
End Function
 
' テーブルデータを全削除する（resetAutoNoにオートナンバー型のカラム名を設定すると自動採番を1にリセット）
Function DeleteTable(ByVal accdbPath As String, ByVal tableName As String, _
                    Optional ByVal resetAutoNo As String = "") As Boolean
    DeleteTable = False
    Application.StatusBar = "「" & tableName & "」テーブルのデータ削除実行..."
    ' エラー処理
    On Error GoTo EXE_ERR
    ' データベース設定
    Dim accCn As New ADODB.Connection
    Dim accRs As New ADODB.Recordset
    accCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accdbPath & ";"
    accRs.CursorLocation = 3
    accRs.Open "SELECT 1 FROM " & tableName, accCn
    ' データ削除判定、実行
    If accRs.RecordCount > 0 Then
        accCn.Execute "DELETE FROM " & tableName
        If resetAutoNo <> "" Then
            accCn.Execute "ALTER TABLE " & tableName & " ALTER COLUMN " & resetAutoNo & " NUMBER"
            accCn.Execute "ALTER TABLE " & tableName & " ALTER COLUMN " & resetAutoNo & " COUNTER (1,1)"
        End If
    End If
CLOSE_DB:
    ' DB切断
    accCn.Close
    Set accCn = Nothing
    DeleteTable = True
    Exit Function
EXE_ERR:
    Call Common.ErrorMessage("DeleteTable", "失敗したテーブル：" & tableName)
End Function
