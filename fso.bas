' ファイルorフォルダチェック
Function IsFileFolder(ByVal path As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then
        IsFileFolder = 1
    ElseIf fso.FileExists(path) Then
        IsFileFolder = 2
    Else
        IsFileFolder = 0
    End If
End Function
  
