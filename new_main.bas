Attribute VB_Name = "Main"
Option Explicit

' #############################################################
'
'
'
'
'
'
' #############################################################

Dim errMsgString As String

Sub シナリオ作成()
    Call executeWithErrorHandling("シナリオ作成")
End Sub


' #############################################################
'
' プライベート変数定義
'
' #############################################################
Private Sub executeWithErrorHandling(processName As String)
    CommonFunction.ProcessStartSettings
    On Error GoTo ErrorHandler
    
    Select Case processName
        Case "シナリオ作成"
            Call convertDecisionToScenario
        Case Else
            Err.Raise 9999, "Main", "未定義"
    End Select
    
    CommonFunction.ProcessEndSettings
    Exit Sub
    
    ' エラー時処理
ErrorHandler:
    CommonFunction.ErrorMessage processName, _
                                "エラー内容：" & Err.Description

End Sub

Private Sub convertDecisionToScenario()
    Dim exeCheck As Integer
    exeCheck = MsgBox("実行確認", vbYesNoCancel + vbInformation)
    Select Case exeCheck
        Case vbYes
        Case vbNo
            Exit Sub
        Case vbCancel
            Exit Sub
    End Select


    ' 使用範囲セル
    Dim usedRangeData As Variant
    usedRangeData = Worksheets("新テーブル").UsedRange
    ' テストシナリオセル
    Dim testScenarioCell As Collection
    Set testScenarioCell = CommonFunction.FindAllStringInSheet("#テストシナリオ", "新テーブル")
    ' テストシナリオ範囲／使用範囲
    Dim testStartRow As Integer: testStartRow = 0
    Dim testStartCol As Integer: testStartCol = 0
    Dim testEndRow As Integer: testEndRow = 0
    Dim testEndCol As Integer: testEndCol = 0
    ' テストシナリオ条件列数（条件列含む）
    Dim jokenColCount As Integer: jokenColCount = 4
    
    If testScenarioCell.Count = 1 Then
        testStartRow = testScenarioCell.Item(1).Row
        testStartCol = testScenarioCell.Item(1).Column
        testEndRow = UBound(usedRangeData)
        testEndCol = UBound(usedRangeData, 2)
    Else
        CommonFunction.ErrorMessage ("不正データ")
    End If
    
    If testStartCol + jokenColCount < testEndCol Then
    
    End If
    
    ' シナリオデータ（デシジョンテーブル結果をテストシナリオ出力用に）
    Dim scenarioData As Variant
    
    
    
    

End Sub
