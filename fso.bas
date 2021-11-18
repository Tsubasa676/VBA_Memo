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
  
Private Function GetPath() As String
    GetPath = ""
    ChDir ThisWorkbook.Path
    GetPath = Application.GetOpenFilename("全てのファイル,*.*")
    If GetPath = "False" Then
        GetPath = ""
        Exit Function
    End If
End Function
        
        
        
