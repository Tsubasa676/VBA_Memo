Option Explicit

Private Sub Workbook_Open()

End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' 範囲チェック
    Dim isRange As Boolean
    ' 空白チェック
    Dim isBlank As Boolean
    ’ テストシート用
    If Sh.Name = VAL_CREATE_SHEET Then
        isRange = Target.Row > 10 And Target.Column > 4
        isBlank = Target.Value = ""
        If isRange Then
            Cancel = True ' 入力状態にしない
            Target.Value = IIf(isBlank, "●", "")
        End If
    End If
    
End Sub
