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



Public Function ErrMsg(ByVal message As String)
    MsgBox message, vbOKOnly + vbExclamation, "エラー"
End Function

Public Function QAMsg(ByVal message As String)
    MsgBox message, vbYesNo + vbQuestion, "確認"
End Function
