Option Explicit

Private RowKa As Long         ' #確認項目　行
Private RowKi As Long         ' #期待値  　行
Private CaseRange As Variant  ' テストケースの対象セル範囲
Private CaseArr() As String

' 作成
Sub Sakusei()
    Application.ScreenUpdating = False
    ' チェック
    If SetVal <> True Then Exit Sub
    ' 変数定義
    Dim LastRow As Long: LastRow = SB_Kitaichi_LastRow(RowKi) - 1
    Dim LastColumn As Long: LastColumn = GetLastColumn(RowKa)
    Dim CaseCnt As Long: CaseCnt = LastColumn - COL_TEST_CASE + 1
    If CaseCnt < 1 Then Exit Sub
    ' テストケース連番振り直し
    Call SB_TestCase_CaseNo(RowKa, LastColumn)
    ' 入力用配列
    CaseRange = Range(Cells(RowKa, COL_JOUKEN_KITAICHI), Cells(LastRow, LastColumn))
    ' 出力用配列
    Erase CaseArr
    ReDim Preserve CaseArr(1 To CaseCnt, 1 To VAL_SCENARIO_CNT)
    
    Dim i As Long, j As Long, c As Long, kgr As Long
    c = 1
    kgr = RowKi - RowKa
    For i = 4 To UBound(CaseRange, 2)
        For j = 1 To UBound(CaseRange)
            If j = 1 Then
                CaseArr(c, 1) = ActiveSheet.Name & "-" & CaseRange(1, i)
                CaseArr(c, 4) = Range(RNG_YOTEIBI)
                CaseArr(c, 5) = Range(RNG_YOTEISYA)
            ElseIf j > 1 And j <= kgr And CaseRange(j, i) <> "" Then
                If "" & CaseRange(j, 2) & CaseRange(j, 3) <> "" Then
                    CaseArr(c, 2) = CaseArr(c, 2) & CaseRange(j, 1) & "[" & CaseRange(j, 2) & "]" & CaseRange(j, 3) & Chr(10)
                Else
                    CaseArr(c, 2) = CaseArr(c, 2) & CaseRange(j, 1) & Chr(10)
                End If
            ElseIf j > 1 And j > kgr And CaseRange(j, i) <> "" Then
                CaseArr(c, 3) = CaseArr(c, 3) & CaseRange(j, 1) & Chr(10)
            End If
        Next
        c = c + 1
    Next
    If Sakusei_Sub <> True Then ErrMsg (VAL_SCENARIO_SHEET & "へ出力できませんでした。")
    Application.ScreenUpdating = True
End Sub
Private Function Sakusei_Sub() As Boolean
    Sakusei_Sub = False
    On Error Resume Next
    Dim EndRow As Long
    With ThisWorkbook.Worksheets(VAL_SCENARIO_SHEET)
        .Rows("2:" & Rows.Count).ClearContents
        .Rows("2:" & Rows.Count).Borders.LineStyle = xlLineStyleNone
        .Rows("2:" & Rows.Count).Interior.Color = RGB(255, 255, 255)
        .Range("A2").Resize(UBound(CaseArr), UBound(CaseArr, 2)) = CaseArr
        EndRow = .Cells(Rows.Count, 1).End(xlUp).Row
        With .Range(.Cells(2, 1), .Cells(EndRow, UBound(CaseArr, 2)))
            .Borders.LineStyle = True
            .WrapText = True
        End With
        With .Range(.Cells(2, 8), .Cells(EndRow, 8)).Validation
            .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="OK,NG,実施不可,不具合"
        End With
    End With
    
    Sakusei_Sub = True
End Function



' 条件、期待値の行を求める
Private Function SetVal() As Boolean
    SetVal = False
    On Error Resume Next
    RowKa = GetRowSearch(COL_JOUKEN_KITAICHI, VAL_KAKUNIN)
    RowKi = GetRowSearch(COL_JOUKEN_KITAICHI, VAL_KITAICHI)
    If (RowKa < 1 Or RowKi < 1) Or (RowKa > RowKi) Then
        Call ErrMsg("設定内容に不整合が生じています。")
        Exit Function
    End If
    SetVal = True
End Function

' 条件
Private Sub SB_Jouken_SpinUp()
    Application.ScreenUpdating = False
    ' チェック
    If SetVal <> True Then Exit Sub
    If (RowKi - RowKa) <= 4 Then Exit Sub
    Rows(RowKi - 1).Delete
    Application.ScreenUpdating = True
End Sub
Private Sub SB_Jouken_SpinDown()
    Application.ScreenUpdating = False
    ' チェック
    If SetVal <> True Then Exit Sub
    Rows(RowKi - 1).Copy
    Rows(RowKi).Insert (xlShiftDown)
    Application.CutCopyMode = False
    Rows(RowKi).Interior.Color = RGB(255, 255, 255)
    Rows(RowKi).ClearContents
    Rows(RowKi).HorizontalAlignment = xlLeft
    Application.ScreenUpdating = True
End Sub


' 期待値
Private Sub SB_Kitaichi_SpinUp()
    Application.ScreenUpdating = False
    ' チェック
    If SetVal <> True Then Exit Sub
    ' 最終行削除
    Dim Row As Long: Row = SB_Kitaichi_LastRow(RowKi)
    If Row = 0 Or Row - 4 <= RowKi Then Exit Sub
    Rows(Row - 1).Delete
    Application.ScreenUpdating = True
End Sub
Private Sub SB_Kitaichi_SpinDown()
    Application.ScreenUpdating = False
    Dim Row As Long: Row = GetLastRow(COL_JOUKEN_KITAICHI)
    If SetVal <> True Then Exit Sub
    Rows(Row).Copy
    Rows(Row + 1).Insert (xlShiftDown)
    Application.CutCopyMode = False
    With Range(Cells(Row + 1, COL_JOUKEN_KITAICHI), Cells(Row + 1, COL_JOUKEN_KITAICHI + 2))
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    Rows(Row + 1).Interior.Color = RGB(255, 255, 255)
    Rows(Row + 1).ClearContents
    Application.ScreenUpdating = True
End Sub
Private Function SB_Kitaichi_LastRow(ByVal RowKi As Long) As Long
    SB_Kitaichi_LastRow = 0
    Dim RowUR As Long: RowUR = UsedRange.Rows.Count
    If RowKi = RowUR Then Exit Function
    Dim i As Long
    For i = RowKi + 1 To RowUR
        If Cells(i, COL_JOUKEN_KITAICHI).Borders(xlEdgeBottom).LineStyle <> xlContinuous Then
            Exit For
        End If
    Next i
    SB_Kitaichi_LastRow = i
End Function


' テストケース
Private Sub SB_TestCase_SpinUp()
    Application.ScreenUpdating = False
    ' チェック
    If SetVal <> True Then Exit Sub
    Dim LastColumn As Long: LastColumn = GetLastColumn(RowKa)
    If LastColumn < COL_TEST_CASE - 1 Then
        Call ErrMsg("設定内容に不整合が生じています。")
        Exit Sub
    End If
    Columns(LastColumn).Copy
    Columns(LastColumn + 1).Insert (xlShiftToRight)
    Columns(LastColumn + 1).ClearContents
    Application.CutCopyMode = False
    Cells(RowKa, LastColumn + 1).Select
    Call SB_TestCase_CaseNo(RowKa, LastColumn + 1)
    Application.ScreenUpdating = True
End Sub
Private Sub SB_TestCase_SpinDown()
    Application.ScreenUpdating = False
    Dim Row As Long
    Dim LastColumn As Long
    Row = GetRowSearch(COL_JOUKEN_KITAICHI, VAL_KAKUNIN)
    LastColumn = GetLastColumn(Row)
    If LastColumn - 2 < COL_TEST_CASE Then Exit Sub
    Columns(LastColumn).Delete
    Cells(Row, LastColumn - 1).Select
    Application.ScreenUpdating = True
End Sub
Private Sub SB_TestCase_CaseNo(ByVal Row As Long, ByVal LastColumn As Long)
    ' 連番振り直し
    Dim i As Long
    Dim c As Long: c = 1
    For i = COL_TEST_CASE To LastColumn
        Columns(i).HorizontalAlignment = xlCenter
        Columns(i).ColumnWidth = 4
        Cells(Row, i) = "c" & c
        c = c + 1
    Next
End Sub


' ダブルクリック時
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim Row As Long: Row = Target.Row
    Dim Col As Long: Col = Target.Column
    Dim RowKa As Long: RowKa = GetRowSearch(COL_JOUKEN_KITAICHI, VAL_KAKUNIN)
    Dim RowKi As Long: RowKi = GetRowSearch(COL_JOUKEN_KITAICHI, VAL_KITAICHI)
    ' チェック
    If RowKa < 1 Or RowKi < 1 Then
        Call ErrMsg("設定内容に不整合が生じています。")
        Exit Sub
    End If
    ' 処理分岐
    If Col = COL_JOUKEN_KITAICHI + 1 And Row > RowKa And Row < RowKi Then
        KakuninKubun.Show
        Cancel = True
    ElseIf Col >= COL_TEST_CASE And Row > RowKa Then
        If Target.Value = "" Then
            Target.Value = "＊"
        Else
            Target.Value = ""
        End If
        Cancel = True
    End If
End Sub







