Attribute VB_Name = "modUtil"
Option Explicit

'===========================================
' modUtil：汎用（＋シート操作＋ファイル操作）
'===========================================

'シート存在チェック
Public Function SheetExists(sheetName As Variant, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet

    '引数 wb が省略された場合は、このブック（ThisWorkbook）を使用
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    On Error Resume Next ' エラーを無視して処理を続行
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0      ' エラー処理を元に戻す

    SheetExists = Not ws Is Nothing
    
End Function



'"シート・列"を指定し、最後の行を取得
Public Function GetLastRow(ws As Worksheet, col As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function



'フィルター解除
Public Sub ClearFilter(ws As Worksheet)
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
End Sub



'ファイルが開いているかチェック
Public Function IsFileOpen(filePath As String) As Boolean
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Binary Access Read Write Lock Read Write As #ff
    Close #ff
    If Err.Number <> 0 Then
        IsFileOpen = True
        Err.Clear
    Else
        IsFileOpen = False
    End If
End Function



'配列に要素を追加する
Public Function AddToArray(arr As Variant, value As Variant) As Variant
    Dim src() As Variant
    Dim dst() As Variant
    Dim i As Long

    '--- SafeArray を通常配列に変換 ---
    src = arr

    '--- 新しい配列を作成 ---
    ReDim dst(LBound(src) To UBound(src) + 1)

    '--- コピー ---
    For i = LBound(src) To UBound(src)
        dst(i) = src(i)
    Next i

    '--- 新しい値を追加 ---
    dst(UBound(dst)) = value

    AddToArray = dst
End Function


