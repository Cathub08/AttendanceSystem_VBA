Attribute VB_Name = "modAttendanceSheet"
Option Explicit

'===================================
' modAttendanceSheet：勤怠シート管理
'===================================

'Thisworkbook全量のAttendanceSheets名のコレクション
Public TgName As Collection


'今月のAttendanceSheets名を取得する
Public Function GetAttendanceSheetName(Optional targetDate As Date) As String
    If targetDate = 0 Then targetDate = Date
    GetAttendanceSheetName = "Attendance_" & Format(targetDate, "yyyy_mm")
End Function


'Thisworkbook全量のAttendanceSheets名を取得する
Public Sub GetAllAttendanceSheetsName()
    Dim wb As Workbook
    Dim i As Long
    Dim keyword As String
    
    keyword = SHEET_ATT_PREFIX
    Set wb = ThisWorkbook
    Set TgName = New Collection
    
    For i = 1 To wb.Sheets.Count
        If InStr(1, wb.Sheets(i).name, keyword, vbTextCompare) > 0 _
            And wb.Sheets(i).name <> SHEET_ATT_TEMPLATE Then
        
            TgName.Add wb.Sheets(i).name
        End If

    Next i
End Sub


'未来3か月分のAttendanceSheets名を取得する
Public Function GetFutureAttendanceSheetsName() As Collection
    Dim d As Date
    Dim ThreeMonthslater As Date
    Dim nameItem As Variant
    Dim targetName As Collection
    
    d = DateSerial(year(Date), month(Date), 1) 'Dateでも問題ないが当月1日にすることで月末に実行したときの “日付のズレ” を完全に排除
    ThreeMonthslater = DateAdd("m", 3, d)
    Set targetName = New Collection
        
    Do While d <= ThreeMonthslater '今月の1日から3か月
        On Error Resume Next
        targetName.Add "Attendance_" & year(d) & "_" & Format(month(d), "00"), "Attendance_" & year(d) & "_" & Format(month(d), "00") '重複値は省いてCollectionに登録。「,」の後は重複できないKeyの指定。[Key:=]の表示をここでは省略している。
        On Error GoTo 0
        d = DateAdd("m", 1, d)
    Loop
    
    Set GetFutureAttendanceSheetsName = targetName
End Function


'今月～3か月分(例:今が5月なら8月分まで作成)Attendanceシートが存在するか、なければシートを追加＆シート名をつける
Public Sub SheetExistsAndGetFutureSheet()
    Dim wb As Workbook
    Dim d As Date
    Dim ThreeMonthslater As Date
    Dim ws As Worksheet
    Dim nameItem As Variant
    Dim found As Boolean
    Dim GetName As Collection
    Dim NotFound As Collection
    Dim AdditionTarget As Variant
     
    Set wb = ThisWorkbook
    d = DateSerial(year(Date), month(Date), 1) 'Dateでも問題ないが当月1日にすることで月末に実行したときの “日付のズレ” を完全に排除
    ThreeMonthslater = DateAdd("m", 3, d)
    Set NotFound = New Collection
    Set GetName = GetFutureAttendanceSheetsName
    
    If Not SheetExists("Attendance_Template") Then Exit Sub
        
    For Each nameItem In GetName
        found = False
    
        For Each ws In ThisWorkbook.Worksheets
            If StrComp(ws.name, nameItem, vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next ws
        
        If found = False Then
        NotFound.Add nameItem
        End If
    Next nameItem
 
    If NotFound.Count > 0 Then
        For Each AdditionTarget In NotFound
            '存在しない場合の処理
            wb.Sheets("Attendance_Template").Copy After:=wb.Sheets(wb.Sheets.Count)
            wb.ActiveSheet.name = AdditionTarget
            With wb.ActiveSheet
                If .AutoFilterMode = False Then 'フィルターが二重に設定される事故を防ぐ
                    .Range("A1").AutoFilter
                End If
            End With
        Next AdditionTarget
    End If
    
End Sub


'データ管理（Attendance シートをアクティブにする）
Public Function Act_Att_sheet() As result
    On Error GoTo ErrHandler
    
    '初期値
    Act_Att_sheet = Other
    
    Dim sheetName As String
    sheetName = GetAttendanceSheetName

    If Not SheetExists(sheetName) Then
            Act_Att_sheet = InvalidInput
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    Sheets(sheetName).Columns("A:K").AutoFit
    
    ws.Activate
    Act_Att_sheet = Success
    Exit Function
    
ErrHandler:
    Act_Att_sheet = Other '想定外の異常
End Function

