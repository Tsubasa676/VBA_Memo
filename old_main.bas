Attribute VB_Name = "MainProcedure"
Option Explicit

' ################# シナリオ作成→テストシナリオ 設定 #######################
Private RowJouken As Long                         ' 条件の開始行
Private RowResult As Long                         ' 結果の開始行

Private Const ColumnFirstCase As Long = 5         ' 最初のケース列番号
Private ColumnLastCase As Long                    ' 最後のケース列番号

Private ArrCaseInput As Variant              ' テストケースの対象セル範囲
Private ArrCaseOutput() As String            ' 作成
Private Type ArrCaseLabelType                ' s
    label As String
    startRow As Integer
    endRow As Integer
    OutputCol As Integer
End Type
Private ArrCaseLabel() As ArrCaseLabelType   ' 列名
' セル内改行出力はChr(10)

' ################# テストシナリオ→各結果シート 設定 #######################
Private Type ScenarioToResult
    InputRow As Integer
    InputCol As Integer
    OutputRow As Integer
    OutputCol As Integer
End Type

' ###########################################################################
Sub 設定を開く()
    MainForm.Show
End Sub


Sub シナリオ作成()
    ProcessStartSettings
    On Error GoTo Exception
    
    ' 処理
    MainProcedure.DecisionToScenario
    
    ProcessEndSettings
    Exit Sub
Exception:
    CommonFunction.ErrorMessage ("シナリオ作成 から テストシナリオ 作成中にエラーが発生しました。")
    ProcessEndSettings
End Sub

Sub シナリオクリア()
    ProcessStartSettings
    On Error GoTo Exception
    
    ' テストシナリオをクリア
    MainProcedure.ScenarioClear
    
    ProcessEndSettings
    Exit Sub
Exception:
    ProcessEndSettings
    CommonFunction.ErrorMessage ("テストシナリオ クリア時にエラーが発生しました。")
End Sub

Sub 結果シート作成()
    ProcessStartSettings
    On Error GoTo Exception
    
    If MsgQA("テストシナリオから結果シートを作成しますか？") <> True Then
        Exit Sub
    End If

    If SheetExists(SHEET_HINAGATA) <> True Then
        MsgErr (SHEET_HINAGATA + "シートが存在しないため実行できません")
        Exit Sub
    End If
    
    Call CommonFunction.SheetShowOrHide(SHEET_HINAGATA, True)
    
    ' シナリオデータ作る
    Dim ws_scenario As Worksheet
    Set ws_scenario = ThisWorkbook.Worksheets(SHEET_TEST_SCENARIO)
    
    ' シートの最終列
    Dim Col_LastOfSheet As Integer: Col_LastOfSheet = GetCol_LastOfSheet(ws_scenario) - SCENARIO_ADD_COL
    ' シートの最終行
    Dim Row_LastOfSheet As Integer: Row_LastOfSheet = GetRow_LastOfSheet(ws_scenario)

    If SCENARIO_START_ROW >= Row_LastOfSheet Or SCENARIO_START_COL >= Col_LastOfSheet Then
        MsgErr ("先に テストシナリオ を作成してください。")
        Exit Sub
    End If

    ' シナリオのデータ格納
    Dim ArrScenarioInput As Variant: _
        ArrScenarioInput = Range(ws_scenario.Cells(SCENARIO_START_ROW, SCENARIO_START_COL) _
                               , ws_scenario.Cells(Row_LastOfSheet, Col_LastOfSheet))
    
    ' シートコピー、シナリオ転記
    Dim ws_hina As Worksheet
    Set ws_hina = ThisWorkbook.Worksheets(SHEET_HINAGATA)
    
    ' ループ番号
    Dim i As Integer, j As Integer, s As Integer
    ' アラートフラグ
    Dim alertFlag As Boolean: alertFlag = False
    ' シート重複フラグ
    Dim sheetExistsFlag As Boolean: sheetExistsFlag = False
    ' シナリオチェック
    Dim isNullScenario As Boolean

    ' 処理
    With ThisWorkbook
        For i = 1 To UBound(ArrScenarioInput)
            If i <> 1 Then
                ' 条件期待値チェック
                isNullScenario = True
                For s = 2 To UBound(ArrScenarioInput, 2)
                    If ArrScenarioInput(i, s) <> "" Then
                        isNullScenario = False
                    End If
                Next
                If isNullScenario = True Then
                    GoTo Continue
                End If
            
                ' シート重複チェック
                sheetExistsFlag = CommonFunction.SheetExists(CStr(i - 1))
                If sheetExistsFlag = True And alertFlag <> True Then
                    alertFlag = CommonFunction.MsgQA("既に同じ名前の結果シートが存在します。" & vbCrLf & "全ての結果シートのシナリオを上書きしますか？")
                    If alertFlag <> True Then
                        CommonFunction.MsgInfo ("結果シート作成を中止しました。")
                        Exit For
                    End If
                End If
                
                ' 雛形シートをコピー
                If sheetExistsFlag <> True Then
                    .Worksheets(SHEET_HINAGATA).Copy After:=.Worksheets(Worksheets.Count)
                    .ActiveSheet.Name = i - 1
                End If
                
                ' シナリオ貼り付け
                With .Worksheets(CStr(i - 1))
                    With .Rows("1:2")
                        .ClearContents
                        .HorizontalAlignment = xlGeneral
                        .Borders.LineStyle = xlLineStyleNone
                        .Interior.Color = RGB(255, 255, 255)
                    End With
                    For j = 1 To UBound(ArrScenarioInput, 2)
                        .Cells(1, j) = ArrScenarioInput(1, j)
                        .Cells(2, j) = ArrScenarioInput(i, j)
                    Next
                    ' カラム体裁
                    With .Range(.Cells(1, 1), .Cells(1, UBound(ArrScenarioInput, 2)))
                        .Interior.Color = RGB(189, 215, 238)
                        .HorizontalAlignment = xlCenter
                    End With
                    ' 表全体体裁
                    With .Range(.Cells(1, 1), .Cells(2, UBound(ArrScenarioInput, 2)))
                        .Borders.LineStyle = True
                        .WrapText = True
                    End With
                End With
            End If
Continue:
        Next
    End With
    WorksheetsSort
    ThisWorkbook.Worksheets(SHEET_SCENARIO_MATRIX).Activate
    ProcessEndSettings
    Exit Sub
Exception:
    CommonFunction.ErrorMessage (SHEET_SCENARIO_MATRIX)
End Sub

' 結果シートのソート
Public Function WorksheetsSort()
    Dim i As Integer, j As Integer
    Dim ws As Worksheet
    Dim numSheets() As String
    Dim sheetCount As Integer
    Dim testScenarioIndex As Integer
    Dim temp As String
    ReDim numSheets(1 To ThisWorkbook.Sheets.Count)
    sheetCount = 0
    testScenarioIndex = 0
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = Constant.SHEET_TEST_SCENARIO Then
            testScenarioIndex = i
        End If
    Next i
    If testScenarioIndex = 0 Then Exit Function
    For Each ws In ThisWorkbook.Sheets
        If IsNumeric(Left(ws.Name, 1)) Then
            sheetCount = sheetCount + 1
            numSheets(sheetCount) = ws.Name
        End If
    Next ws
    ReDim Preserve numSheets(1 To sheetCount)
    For i = 1 To sheetCount - 1
        For j = 1 To sheetCount - i
            If Val(numSheets(j)) > Val(numSheets(j + 1)) Then
                temp = numSheets(j)
                numSheets(j) = numSheets(j + 1)
                numSheets(j + 1) = temp
            End If
        Next j
    Next i
    For i = 1 To sheetCount
        ThisWorkbook.Sheets(numSheets(i)).Move After:=ThisWorkbook.Sheets(testScenarioIndex + i)
    Next i
End Function

' シナリオ作成からテストシナリオへ
Public Function DecisionToScenario() As Boolean
    DecisionToScenario = False
    
    ' 実行確認
    Dim rc As Integer
    rc = MsgBox(SHEET_SCENARIO_MATRIX + " シートを" + SHEET_TEST_SCENARIO + " シートに転記します。" + vbCrLf _
              + SHEET_TEST_SCENARIO + " シートをクリアしますか？" + vbCrLf + vbCrLf _
              + "　はい：クリア（確認）して実行" + vbCrLf + "　いいえ：クリアせず実行（実施者等の列はクリア）" + vbCrLf + "　キャンセル：実行中止" _
              , vbYesNoCancel + vbInformation, "実行確認")
    
    Select Case rc
        Case vbYes
            MainProcedure.ScenarioClear
        Case vbCancel
            Exit Function
    End Select
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SCENARIO_MATRIX)
    
    ' 「条件」開始行
    Dim RowJouken As Integer: RowJouken = GetRow_Search(ws, COLUMN_START, VAL_JOUKEN)
    ' 「結果（期待値）」開始行
    Dim RowResult As Integer: RowResult = GetRow_Search(ws, COLUMN_START, VAL_RESULT)
    ' シートの最終列
    Dim Col_LastOfSheet As Integer: Col_LastOfSheet = GetCol_LastOfSheet(ws)
    ' シートの最終行
    Dim Row_LastOfSheet As Integer: Row_LastOfSheet = GetRow_LastOfSheet(ws)
        
    ' テストケース連番振り直し
    Call SB_TestCase_CaseNo(ws, RowJouken, COLUMN_START_MARK, Col_LastOfSheet)
    Call SB_TestCase_CaseNo(ws, RowResult, COLUMN_START_MARK, Col_LastOfSheet)
    
    ' シナリオのデータ格納
    Dim ArrScenarioInput As Variant: _
        ArrScenarioInput = Range(ws.Cells(RowJouken, COLUMN_START) _
                               , ws.Cells(Row_LastOfSheet, Col_LastOfSheet))
    
    ' シナリオのデータから列名を格納する
    ReDim ArrScenarioColumn(1 To 4, 1 To 1) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = LBound(ArrScenarioInput) To UBound(ArrScenarioInput)
        If ArrScenarioInput(i, 1) <> "" And ArrScenarioInput(i, 1) <> VAL_RESULT Then
            ' 条件行の設定
            If GetCol_LastOfSpecifiedRow(ws, RowJouken + i - 1) >= COLUMN_START_MARK Then
                j = j + 1
                ReDim Preserve ArrScenarioColumn(1 To UBound(ArrScenarioColumn, 1), 1 To j)
                If i = 1 Then
                    ArrScenarioColumn(1, j) = "No." ' 列名
                Else
                    ArrScenarioColumn(1, j) = ArrScenarioInput(i, 1)
                    
                End If
                ArrScenarioColumn(2, j) = i         ' 入力開始行
                ArrScenarioColumn(4, j) = j         ' 出力列
            End If
            
            ' 各条件、結果の入力終了行格納
            For k = i To UBound(ArrScenarioInput)
                If k <> i And ArrScenarioInput(k, 1) <> "" And IsEmpty(ArrScenarioColumn(3, j)) Then
                    ArrScenarioColumn(3, j) = k - 1 ' 入力終了行
                    Exit For
                End If
            Next
        End If
        
        ' 最後の結果列
        If i = UBound(ArrScenarioInput) Then
            ArrScenarioColumn(3, j) = i
        End If
    Next
    ' シナリオチェック
    If j < 3 Then
        MsgErr ("条件、結果を設定してください")
        Exit Function
    End If
    
    ' シナリオシート：追加列
    Dim AddScenarioColumn(1 To SCENARIO_ADD_COL) As String
    AddScenarioColumn(1) = "実施者"
    AddScenarioColumn(2) = "実施日"
    AddScenarioColumn(3) = "テスト結果"
    AddScenarioColumn(4) = "備考"

    ' シナリオに列を追加
    ReDim Preserve ArrScenarioColumn(1 To UBound(ArrScenarioColumn, 1), 1 To j + UBound(AddScenarioColumn, 1))
    For i = 1 To UBound(AddScenarioColumn)
        ArrScenarioColumn(1, j + i) = AddScenarioColumn(i)
    Next
    
    ' シナリオ出力準備
    Dim ArrScenarioOutput() As Variant
    ReDim Preserve ArrScenarioOutput(1 To Col_LastOfSheet - COLUMN_START_MARK + 1, 1 To UBound(ArrScenarioColumn, 2))
    Dim RowOutput As Long: RowOutput = 0
    Dim ColOutput As Long: ColOutput = 1
    ' シナリオ出力用の列ループ
    For i = 1 To UBound(ArrScenarioInput, 2)
    
        ' シナリオ開始列までスキップ
        If i > COLUMN_START_MARK - COLUMN_START Then
        
            ' シナリオ出力行カウント
            RowOutput = RowOutput + 1
            
            ' シナリオの行ループ
            For j = 1 To UBound(ArrScenarioInput)
            
                If j = 1 Then
                    ' シナリオのNoを設定
                    ArrScenarioOutput(RowOutput, 1) = ArrScenarioInput(j, i)
                ElseIf ArrScenarioInput(j, i) <> "" Then
                    ' 初期化
                    ColOutput = 0
                    ' シナリオがどの条件になるのか取得
                    For k = 1 To UBound(ArrScenarioColumn, 2)
                        ' 各条件開始行 <= j And j <= 各条件終了行
                        If ArrScenarioColumn(2, k) <= j And j <= ArrScenarioColumn(3, k) Then
                            ' 出力列
                            ColOutput = ArrScenarioColumn(4, k)
                        End If
                    Next
                    
                    ' 条件に一致するか
                    If ColOutput <> 0 Then
                        ' 条件内容を設定
                        If ArrScenarioOutput(RowOutput, ColOutput) <> "" Then
                            ArrScenarioOutput(RowOutput, ColOutput) = ArrScenarioOutput(RowOutput, ColOutput) & Chr(10) _
                                                                    & ArrScenarioInput(j, 2) & ArrScenarioInput(j, 3)
                        Else
                            ArrScenarioOutput(RowOutput, ColOutput) = ArrScenarioInput(j, 2) & ArrScenarioInput(j, 3) & ArrScenarioInput(j, 4)
                        End If
                    End If
                End If
                
            Next
            
        End If
        
    Next
    
    ' シートにシナリオ出力
    Dim RowOutputStart As Integer: RowOutputStart = SCENARIO_START_ROW
    Dim ColOutputStart As Integer: ColOutputStart = SCENARIO_START_COL
    Dim wss As Worksheet
    Set wss = ThisWorkbook.Worksheets(SHEET_TEST_SCENARIO)
    With wss
        ' シナリオ貼り付け
        .Cells(RowOutputStart, ColOutputStart).Resize(UBound(ArrScenarioColumn), UBound(ArrScenarioColumn, 2)) = ArrScenarioColumn
        .Cells(RowOutputStart + 1, ColOutputStart).Resize(UBound(ArrScenarioOutput), UBound(ArrScenarioOutput, 2)) = ArrScenarioOutput
        ' カラム体裁
        With .Range(.Cells(RowOutputStart, ColOutputStart), .Cells(RowOutputStart, UBound(ArrScenarioColumn, 2) + 1))
            .Interior.Color = RGB(189, 215, 238)
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlCenter
        End With
        ' 表全体体裁
        With .Range(.Cells(RowOutputStart, ColOutputStart), .Cells(GetRow_LastOfSheet(wss), UBound(ArrScenarioColumn, 2) + 1))
            .Borders.LineStyle = True
            .VerticalAlignment = xlTop
            .Validation.Delete
            .WrapText = True
            .Font.Size = 11
        End With
        ' 追加カラム体裁 プルダウン
        Dim ColPullDown As Integer: ColPullDown = UBound(ArrScenarioColumn, 2) + ColOutputStart - 2
        With .Range(.Cells(RowOutputStart + 1, ColPullDown), .Cells(GetRow_LastOfSheet(wss), ColPullDown))
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="○,×,△,ー"
            .Font.Size = 20
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="×"
            ' 背景色を設定（例：薄い赤色）
            With .FormatConditions(.FormatConditions.Count)
                .Interior.Color = RGB(255, 153, 153)
            End With
        End With
        ' 追加カラム体裁　実施者,日
        With .Range(.Cells(RowOutputStart + 1, ColPullDown - 2), .Cells(GetRow_LastOfSheet(wss), ColPullDown))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End With
    
    ' 成功
    DecisionToScenario = True
End Function


' シナリオシートのフォーマット
Public Function ScenarioClear()
    ' 実行確認
    If MsgBox(SHEET_TEST_SCENARIO + "シート をフォーマットしますか？", vbYesNo + vbInformation, "実行確認") <> vbYes Then
        Exit Function
    End If

    ' 設定
    Dim RowOutputStart As Integer: RowOutputStart = SCENARIO_START_ROW
    Dim ColOutputStart As Integer: ColOutputStart = SCENARIO_START_COL
    
    With ThisWorkbook.Worksheets(SHEET_TEST_SCENARIO)
        ' 列幅
        .Columns.ColumnWidth = 20
        .Columns(1).ColumnWidth = 1.9
        .Columns(2).ColumnWidth = 6
    
        ' フォーマット
        With .Rows(RowOutputStart & ":" & Rows.Count)
            ' 入力規則 削除
            .Validation.Delete
            ' 書式設定 削除
            .FormatConditions.Delete
            ' 値 削除
            .ClearContents
            ' 左揃え
            .HorizontalAlignment = xlGeneral
            ' 罫線なし
            .Borders.LineStyle = xlLineStyleNone
            ' 背景色：白
            .Interior.Color = RGB(255, 255, 255)
            ' フォントサイズ11
            .Font.Size = 11
        End With
    End With

End Function

Sub A1保存()
    If CommonFunction.MsgQA("全てのシートを「A1」が選択された状態にして保存しますか？") <> True Then
        Exit Sub
    End If
    ProcessStartSettings
    On Error GoTo Exception
    Dim wss As Worksheet
    With ThisWorkbook
        For Each wss In .Worksheets
            If wss.Visible = xlSheetVisible Then
                Application.GoTo reference:=wss.Range("A1"), Scroll:=True
            End If
        Next
        .Worksheets(1).Activate
        .Save
    End With
    ProcessEndSettings
    Exit Sub
Exception:
    CommonFunction.ErrorMessage ("A1保存中にエラーが発生しました。")
    ProcessEndSettings
End Sub

' マクロなしブック保存
Sub マクロなしブック複製()
    If Right(ThisWorkbook.Name, 4) <> "xlsm" Then
        Exit Sub
    End If
    If CommonFunction.MsgQA("マクロなし（.xlsx）ブックを作成しますか？") <> True Then
        Exit Sub
    End If
    
    WorkbookCreate
End Sub

Function WorkbookCreate()
    Dim folderPath As String
    Dim bookName As String
    Dim bookFullName As String
    Dim inputStr  As String
    Dim savePath As String
    Dim timeStamp As String
    timeStamp = Format(Now(), "YYYYMMDD_hhmmss")
    folderPath = ThisWorkbook.Path
    bookFullName = ThisWorkbook.FullName
    bookName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_" & timeStamp & ".xlsx"
    
    If CommonFunction.IsMacOS = True Then
        savePath = folderPath & "/" & bookName
    Else
        savePath = Left(bookFullName, Len(bookFullName) - 5) & "_" & timeStamp & ".xlsx"
    End If
    
    CommonFunction.MsgInfo ("本ブックがマクロなしブック（" & bookName & "）に切り替わります。" & vbCrLf _
    & "ファイルアクセス許可を求められた場合、本ブックが格納されたフォルダに対してアクセス権を付与してください｡ ")
    
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    CommonFunction.MsgInfo ("マクロ無しブックが作成、切り替わりました。")
End Function

Sub 結果シート削除()
    Dim ws As Worksheet
    Dim i As Long
    If CommonFunction.MsgQA("結果シートを削除しますか？" & vbCrLf & "※数値名のシートが削除されます") <> True Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        'シート名が数字かどうかチェック
        If IsNumeric(ws.Name) Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub
