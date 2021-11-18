' ★注意点
' ・数式の挙動とvbaの挙動が異なる場合があるため注意する

' ★Option系
' 変数宣言を強制
Option Explicit
' 配列の開始数値を指定（デフォルトは0）、基本的には指定しない
Option Base 0

' ★参照設定系（初期参照設定以外）


' ★定数の宣言
Const 定数名 As Long = 1

' ActiveWorkbook
' ActiveSheet
' ActiveCell
' vbCrLf
' Application.ScreenUpdating = False
' Application.ScreenUpdating = True

Sub MemoFileSystemObject()
    ' ★FileSystemObjectについて
    ' 参照設定→Microsoft Scripting Runtime
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' エラー時の処理
    On Error Resume Next
End Sub

Sub MemoHairetsu()
    ' ★配列について
    ' セル内のデータを操作する場合は配列に格納して処理、
    ' 結果を配列でまとめて一括出力が早い。
    ' 場合によっては、Collectionの使用も検討する。
    ' 普通の宣言
    Dim Arr0() As String
    ' 配列の大きさを指定（後でサイズ変更できない）
    Dim Arr1(2) As String
    ' 配列の開始～終了番号を指定する宣言
    Dim Arr2(1 To 10) As String
    ' 配列の大きさを変更（既存のデータはクリアされる）
    ReDim Arr0(0 To 1)
    ' 配列の大きさを変更（既存のデータはクリアされない）
    ReDim Preserve Arr0(0 To 10)
    ' 配列をクリア
    Erase Arr2
    ' セルを配列に格納する方法（※配列初期値は1スタートになる）
    ' ※十数万行の2次元配列などを配列でやると、動作はするが数GBのメモリを使用する。
    Dim cellArr() As Variant
    cellArr = Range("A1:c10")
    ' 配列の要素
    Debug.Print "1次元配列　：初期値：" & LBound(Arr0) & "要素数："; UBound(Arr0)
    ' 配列の要素（2次元配列は横の要素も取れる）
    Debug.Print "2次元配列縦：初期値：" & LBound(cellArr, 1) & "要素数："; UBound(cellArr, 1)
    Debug.Print "2次元配列横：初期値：" & LBound(cellArr, 2) & "要素数："; UBound(cellArr, 2)
End Sub

Sub MemoClipboard()
    ' ★クリップボードのデータの取扱について
    ' 参照設定→Microsoft Forms 2.0 Object Library
    Dim buf As String
    Dim buf2 As String
    Dim CB As New DataObject
    buf = "テストメモ"
    With CB
        .SetText buf        '変数のデータをDataObjectに格納する
        .PutInClipboard     'DataObjectのデータをクリップボードに格納する
        .GetFromClipboard   'クリップボードからDataObjectにデータを取得する
        buf2 = .GetText     'DataObjectのデータを変数に取得する
    End With
    MsgBox buf2
End Sub
