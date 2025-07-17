Option Explicit

' 検索文字列の行番号取得
Public Function GetRowSearch(ByVal Col As Long, ByVal word As String) As Long
    GetRowSearch = 0
    Dim EndRow As Long: EndRow = GetLastRow(Col)
    On Error Resume Next
    GetRowSearch = Range(Cells(1, Col), Cells(EndRow, Col)).Find( _
                    What:=word, LookIn:=xlValues, lookat:=xlWhole).Row
End Function

' 最終行取得
Public Function GetLastRow(ByVal Col As Long) As Long
    GetLastRow = 1
    On Error Resume Next
    GetLastRow = Cells(Rows.Count, Col).End(xlUp).Row
End Function

' 最終列取得
Public Function GetLastColumn(ByVal Row As Long) As Long
    GetLastColumn = 1
    On Error Resume Next
    GetLastColumn = Cells(Row, Columns.Count).End(xlToLeft).Column
End Function

' 罫線最終行取得
Public Function GetRuledLastRow(ByVal Col As Long) As Long
    GetRuledLastRow = 1
    On Error Resume Next
    If Cells(10, Col).Borders(xlEdgeBottom).LineStyle = xlContinuous Then
       MsgBox "下線あり"
    Else
       MsgBox "下線なし"
    End If
End Function

' マクロ呼び出し時に選択したセルに選択したファイル名を記載する
Sub SetFileNameToSelectedCells()
    ' 選択中のセルのアドレス取得
    Dim ac As String
    ac = Replace(ActiveCell.Address, "$", "")
    Application.StatusBar = "「" & ac & "」セルに出力します"
    ' 選択したファイルフルパス取得
    Dim fileName As String
    fileName = Common.GetFullPath
    ' ファイルを選択しなかった場合、終了
    If fileName = "" Then GoTo EXE_END
    ' フルパスからファイル名取得
    fileName = Common.GetFileName(fileName)
    With ThisWorkbook.ActiveSheet
        If .Range(ac).value <> "" Then
            If MsgBox("セル「" & ac & "」の値を下記の通り置き換えますか？" & vbCrLf & vbCrLf & vbCrLf _
& .Range(ac).value & vbCrLf _
& "　↓" & vbCrLf _
& fileName & vbCrLf _
                        , vbInformation + vbYesNo, "置換確認") <> vbYes Then GoTo EXE_END
        End If
        ' 選択セルに選択ファイル名を出力
        .Range(ac).value = fileName
    End With
EXE_END:
    Application.StatusBar = False
End Sub

' 存在チェック⇒ファイル：1、フォルダ：2、存在なし：0　を返却
Function NumFileFolder(ByVal path As String) As Integer
    Dim fso As New Scripting.FileSystemObject
    If fso.FileExists(path) Then
        ' ファイルの場合
        NumFileFolder = 1
    ElseIf fso.FolderExists(path) Then
        ' フォルダの場合
        NumFileFolder = 2
    Else
        ' 上記以外の場合
        NumFileFolder = 0
    End If
End Function
 
' 既にファイルを開いているかチェック⇒開いている：True、開いていない：False　を返却
Function IsOpenFile(ByVal path As String) As Boolean
    IsOpenFile = False
    If NumFileFolder(path) <> 1 Then Exit Function
    On Error Resume Next
    Open path For Append As #1
    Close #1
    If ERR.Number > 0 Then IsOpenFile = True
End Function
 
' 選択したファイルのフルパスを返却（未選択・キャンセル時はブランクを返却）
Function GetFullPath(Optional ff As String = "ファイル,*.*", Optional tt As String = "ファイル選択後、開くを押下してください") As String
    ChDir ThisWorkbook.path
    GetFullPath = Application.GetOpenFilename(FileFilter:=ff, MultiSelect:=False, Title:=tt)
    If GetFullPath = "False" Then GetFullPath = ""
End Function
 
' フルパスからファイル名を返却
Function GetFileName(ByVal fullPath As String) As String
    If NumFileFolder(fullPath) <> 1 Then
        GetFileName = "(null)"
        Exit Function
    End If
    Dim fso As New Scripting.FileSystemObject
    GetFileName = fso.GetFileName(fullPath)
End Function
 
' フルパスからファイル名（拡張子抜き）を返却
Function GetFileBaseName(ByVal fullPath As String) As String
    Dim fso As New Scripting.FileSystemObject
    GetFileBaseName = fso.GetBaseName(fullPath)
End Function
 
' フルパスからファイルのディレクトリを返却
Function GetFolderPath(ByVal fullPath As String) As String
    Dim fso As New Scripting.FileSystemObject
    GetFolderPath = fso.GetParentFolderName(fullPath) & "\"
End Function
 
' ファイルコピー（フルパス指定）
Function canCopyFile(ByVal src As String, ByVal dst As String, Optional ByVal overwrite As Boolean = False) As Boolean
    canCopyFile = False
    On Error GoTo ERR
    Dim fso As New Scripting.FileSystemObject
    fso.CopyFile src, dst, overwrite
    Set fso = Nothing
    canCopyFile = True
ERR:
End Function
 
' 存在するシート名かを返却
Function isExistWorksheet(ByVal sheetName As String) As Boolean
    isExistWorksheet = False
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = sheetName Then isExistWorksheet = True
    Next
End Function
 
' 実行エラー時に表示するダイアログ設定
Function ErrorMessage(ByVal methodName As String, Optional ByVal supportedMessage As String = "")
    Dim str As String
    str = "▲下記エラーが発生しました▲" & vbCrLf
    str = str & vbCrLf & "エラーNo　：" & ERR.Number & vbCrLf
    str = str & vbCrLf & "エラー出所：" & ERR.Source & vbCrLf
    str = str & vbCrLf & "エラー詳細：" & ERR.Description & vbCrLf
    If supportedMessage <> "" Then
        str = str & vbCrLf & "---" & vbCrLf & supportedMessage & vbCrLf & "---"
    End If
    MsgBox str, vbOKOnly + vbExclamation, methodName & "実行エラー"
End Function
 
' 検索ワードのセル位置（A1）を返却、Optional設定でOffset指定可
Function GetValueRange(ByVal searchWord As String, Optional ByVal offsetR As Integer = 0, Optional ByVal offsetC As Integer = 0) As String
    GetValueRange = ""
    With ThisWorkbook.ActiveSheet
        Dim rng As Range
        Set rng = .Cells.Find(searchWord, LookIn:=xlValues, lookat:=xlWhole)
        If Not rng Is Nothing Then GetValueRange = rng.offset(offsetR, offsetC).Address(rowabsolute:=False, columnabsolute:=False)
    End With
End Function
 
' 検索ワードのセル位置の文字列を返却、Optional設定でOffset指定可
Function GetValueRangeVal(ByVal searchWord As String, Optional ByVal offsetR As Integer = 0, Optional ByVal offsetC As Integer = 0) As String
    GetValueRangeVal = ""
    With ThisWorkbook.ActiveSheet
        Dim rng As Range
        Set rng = .Cells.Find(searchWord, LookIn:=xlValues, lookat:=xlWhole)
        If Not rng Is Nothing Then
            GetValueRangeVal = rng.offset(offsetR, offsetC).value
        End If
    End With
End Function
 
' 検索ワードの行番号を返却、Optional設定で行番号±N設定可
Function GetValueRow(ByVal searchWord As String, Optional ByVal offset As Integer = 0) As Integer
    GetValueRow = 0
    With ThisWorkbook.ActiveSheet
        Dim rng As Range
        Set rng = .Cells.Find(searchWord, LookIn:=xlValues, lookat:=xlWhole)
        If Not rng Is Nothing Then GetValueRow = rng.Row + offset
    End With
End Function
 
' 検索ワードの列番号を返却、Optional設定で行番号±N設定可
Function GetValueColumn(ByVal searchWord As String, Optional ByVal offset As Integer = 0) As Integer
    GetValueColumn = 0
    With ThisWorkbook.ActiveSheet
        Dim rng As Range
        Set rng = .Cells.Find(searchWord, LookIn:=xlValues, lookat:=xlWhole)
        If Not rng Is Nothing Then GetValueColumn = rng.Column + offset
    End With
End Function
 
' 配列データをエクセルファイルに出力する（既にファイルを開いている場合は異常終了します）
' 出力成功時ファイル名、失敗時ブランクを返却
Function OutputFromArrayToExcel(ByVal outputFilePath As String, ByRef arrCol As Variant, ByRef arrData As Variant, ByVal changeCol As String) As String
    OutputFromArrayToExcel = ""
    On Error GoTo EXE_ERR
    ' 重複ファイルチェック
    If NumFileFolder(outputFilePath) = 1 Then
        Dim fn As String: fn = GetFileName(outputFilePath)
        If MsgBox("「" & fn & "」は既に存在します。" & vbCrLf _
& "上書きしますか？" & vbCrLf & vbCrLf _
& "「いいえ」を押下で別名保存ができます。", vbInformation + vbYesNo, "上書き確認") = vbYes Then GoTo FN_OK
 
    ' ファイルのリネーム処理
RE_FN:
        fn = Application.InputBox(Prompt:="「" & GetFileName(outputFilePath) & "」と異なるファイル名、" & vbCrLf _
& "存在していないファイル名を入力してください。" & vbCrLf & vbCrLf _
& "※拡張子は「.xlsx」を設定してください。", Title:="ファイル名重複", Default:=fn, Type:=2)
        If fn = "" Then GoTo RE_FN
        If Right(fn, 5) <> ".xlsx" Then GoTo RE_FN
        If NumFileFolder(GetFolderPath(outputFilePath) & fn) = 1 Then GoTo RE_FN
        If GetFileName(outputFilePath) <> fn Then
            If MsgBox("どちらのファイルを複製するか選択してください。" & vbCrLf _
& "はい　：" & GetFileName(outputFilePath) & vbCrLf _
& "いいえ：" & fn & vbCrLf _
                        , vbInformation + vbYesNo, "複製ファイル選択") <> vbYes Then
                outputFilePath = GetFolderPath(outputFilePath) & fn
            End If
        End If
    End If
    ' ファイル出力ができる場合の処理
FN_OK:
    Dim newBook As Workbook
    Set newBook = Workbooks.Add
    Dim cArrCol As Variant
    cArrCol = Split(changeCol, ",")
    Dim cArrColRs As Variant
    With newBook
        Dim i As Integer
        For i = LBound(arrCol, 2) To UBound(arrCol, 2)
            .Worksheets(1).Cells(1, i) = arrCol(1, i)
            cArrColRs = Filter(cArrCol, arrCol(1, i))
            If UBound(cArrColRs) <> -1 Or changeCol = "*" Then
                .Worksheets(1).Columns(i).NumberFormatLocal = "@"
            End If
        Next
        If Not IsEmpty(arrData) Then
            .Worksheets(1).Cells(2, 1).Resize(UBound(arrData), UBound(arrData, 2)) = arrData
        End If
        .Worksheets(1).Cells.EntireColumn.AutoFit
 
        Application.DisplayAlerts = False
        .SaveAs outputFilePath
        Application.DisplayAlerts = True
        .Close
    End With
    OutputFromArrayToExcel = GetFileName(outputFilePath)
    Exit Function
EXE_ERR:
End Function

Public Function ErrMsg(ByVal message As String)
    MsgBox message, vbOKOnly + vbExclamation, "エラー"
End Function

Public Function QAMsg(ByVal message As String)
    MsgBox message, vbYesNo + vbQuestion, "確認"
End Function
