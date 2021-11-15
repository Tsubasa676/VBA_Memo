Option Explicit

Public Sub 左上上書き保存()
    Dim wss As Worksheet
    Dim aws As String
    Dim bnk As Long
    bnk = MsgBox("最初のシートをアクティブにして" & vbCrLf & _
                 "上書き保存しますか？", vbYesNo + vbInformation, "左上上書き確認")
    With ActiveWorkbook
        aws = .ActiveSheet.Name
        For Each wss In .Worksheets
            If wss.Visible = xlSheetVisible Then
                Application.Goto Reference:=wss.Range("A1"), Scroll:=True
            End If
        Next
        If bnk = 6 Then
            .Worksheets(1).Activate
        Else
            .Worksheets(aws).Activate
        End If
        .Save
    End With
End Sub
