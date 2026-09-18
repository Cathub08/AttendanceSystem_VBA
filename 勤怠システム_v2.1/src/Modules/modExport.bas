Attribute VB_Name = "modExport"
Option Explicit

'====================
' modExport：CSV出力
'====================


'月次レポートをCSV出力
Public Sub ExportMonthlyReportCSV()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim r As Range
    Dim reportName As String
    Dim arr As Variant
    Dim line As String
    Dim i As Long
    Dim cellValue As Variant
    Dim folderPath As String
    Dim fileName As String

'    'TEST↓----------------------------------------------------
'    Debug.Print "【No13-2】ExportMonthlyReportCSV に到達"
'    Stop 'イミディエイトウィンドウを確認
'    'TEST↑----------------------------------------------------

    reportName = "MonthlyReport_" & Format(Date, "yyyymm")
    
    'シート存在チェック
    If Not SheetExists(reportName) Then
        MsgBox "今月の勤怠レポートがまだ作成されていないため、CSVは出力できません。" _
        & vbCrLf & "月次レポートを作成してから、もう一度お試しください。", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(reportName)

    '保存先を選択
    filePath = Application.GetSaveAsFilename( _
            InitialFileName:="MonthlyReport_" & Format(Date, "yyyymm") & ".csv", _
            FileFilter:="CSVファイル (*.csv), *.csv")
            
    'キャンセル選択時は処理終了
    If filePath = "False" Then Exit Sub
                   
    folderPath = Left(filePath, InStrRev(filePath, "\"))
    
    '保存先が存在しないエラー回避
    If Not FolderExists_FSO(folderPath) Then
        MsgBox "指定したフォルダが存在しません。フォルダを確認してください。", vbExclamation
        Exit Sub
    End If
                                     
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1) 'Mid文字数省略可能→開始位置以降の全文字取得
    
    'ファイル名が空 → エラー回避
    If Trim(fileName) = "" Then
        MsgBox "ファイル名が空のため、保存できません。ファイル名を指定してください。", vbExclamation
        Exit Sub
    End If

    '絵文字(サロゲートペア)有無確認
    If IsSurrogatePair(filePath) Then
        MsgBox "保存先のフォルダ名に絵文字などの特殊文字が含まれているため、CSVを出力できません。" _
               & vbCrLf & "フォルダ名を変更してから、もう一度お試しください。", vbExclamation
        Exit Sub
    End If
                                      
    'CSV保存先有効確認
    If Not IsValidLocation(filePath) Then
        MsgBox "このフォルダはシステム領域のため、CSVを保存できません。" _
               & vbCrLf & "別のフォルダを選択してください。", vbExclamation
        Exit Sub
    End If
                                      
    If IsFileOpen(filePath) Then
        MsgBox "同名のCSVファイルが開かれています。閉じてから再度実行してください。", vbExclamation
        Exit Sub
    End If
            
    'FilesystemObjectのインスタンスを作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '新しいテキストファイルを作成
    Set ts = fso.CreateTextFile(filePath, True, False)

    'ws.UsedRange.Rows→ワークシートで使用されているセル範囲の行
    For Each r In ws.UsedRange.Rows
        'lineを初期化
        line = ""
    
        '空行なら出力せず次の行へスキップ
        If WorksheetFunction.CountA(r) = 0 Then
            line = ""
            GoTo SkipLine
        End If
       
        '行のデータを丸ごと配列で受け取る
        arr = r.value
        
        'UBound(arr, 2)→配列arrの2つ目の次元で使用できる最大のインデックス(配列arrの最大列数)を取得
        For i = 1 To UBound(arr, 2)
        
            'ERROR値「#VALUE」等なら
            If IsError(arr(1, i)) Then
                cellValue = ""
            Else
                cellValue = arr(1, i)
            End If
            
            'セル内カンマは区切り文字としての認識回避
            If InStr(CStr(cellValue), ",") > 0 Then
                cellValue = """" & cellValue & """"
            End If

            line = line & cellValue
            
            If i < UBound(arr, 2) Then line = line & "," 'カンマでつなぐ
        Next i

        'テキスト(line)を書き込み
        ts.WriteLine line
SkipLine:
    Next r

    'ファイルを閉じる
    ts.Close
    MsgBox "CSVを出力しました。", vbInformation
    
End Sub


'【共通関数】サロゲートペア(絵文字)かどうか判定
Public Function IsSurrogatePair(filePath As String) As Boolean
    Dim i As Long
    Dim code As Long
    
    For i = 1 To Len(filePath)
        code = AscW(Mid(filePath, i, 1)) '→ファイルパスの文字の1文字(開始位置iから1文字)のAscW(引数の先頭の文字のUnicodeコードポイントを返す)をcodeに入れる
        
        '上位サロゲート（絵文字の開始領域）
        If code >= &HD800 And code <= &HDBFF Then 'ハイサロゲート(U+D800～U+DBFF)→VB表記：&HD800～&HDBFF
            IsSurrogatePair = True  '← 絵文字を見つけたら True
            Exit Function
        End If
    Next i
    
    IsSurrogatePair = False   '← 絵文字なし

End Function



'【共通関数】CSV保存先は有効な場所か
Public Function IsValidLocation(filePath As String) As Boolean
    Dim lowerPath As String
    lowerPath = LCase(filePath) '小文字で比較
    
        If InStr(lowerPath, "c:\windows") > 0 _
           Or InStr(lowerPath, "c:\program files") > 0 _
           Or InStr(lowerPath, "c:\program files (x86)") > 0 _
           Or InStr(lowerPath, "c:\$recycle.bin") > 0 _
           Or InStr(lowerPath, "c:\recovery") > 0 _
           Or InStr(lowerPath, "c:\boot") > 0 _
           Or InStr(lowerPath, "c:\perflogs") > 0 Then
              
            IsValidLocation = False '上記に該当したら有効ではない保存先のため、Falseを返す
            Exit Function
        End If
        
    IsValidLocation = True '保存先を有効で返す
    
End Function



'【共通関数】保存先が存在しないエラー回避
Public Function FolderExists_FSO(ByVal folderPath As String) As Boolean
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    FolderExists_FSO = fso.FolderExists(folderPath) 'FolderExists メソッド返却値：フォルダが存在→True。存在なし→False
    
End Function





