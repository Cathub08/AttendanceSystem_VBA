Attribute VB_Name = "modDateUtil"
Option Explicit

'================================
' modDateUtil：勤務時間・残業計算
'================================

'作業時間計算
Public Sub CalcWorkTime(ws As Worksheet, rowNum As Long)

    Dim cin As Date, cout As Date, BREAK_HOURS As Date
    Dim workHours As Double
    Dim overHours As Double
    cin = ws.Cells(rowNum, ATT_COL_CLOCK_IN).value
    cout = ws.Cells(rowNum, ATT_COL_CLOCK_OUT).value
    BREAK_HOURS = ws.Cells(rowNum, ATT_COL_BREAK_TIME).value

    ' --- 単体テスト用コード ↓---
    'cin = CDate("2026/07/1 09:00")
    'cout = CDate("2026/07/1 12:00")
    'BREAK_HOURS = CDate("5:00:00")
    ' --- ここまでテスト用 ↑---

    If IsDate(cin) And IsDate(cout) Then
        Dim wh As Double
        wh = (cout - cin) * 24 - (BREAK_HOURS * 24)
        
        If wh < 0 Then
            wh = 0
        End If
        
        If wh > STD_WORK_HOURS Then '所定労働時間(8h)以上なら
            workHours = STD_WORK_HOURS
            overHours = wh - STD_WORK_HOURS
        Else '所定労働時間未満なら
            workHours = wh
            overHours = 0
        End If
        
    End If
    workHours = Round(workHours * 4, 0) / 4
    overHours = Round(overHours * 4, 0) / 4
           
    ' --- 単体テスト用コード ↓---
    '    Debug.Print "回帰テストNo.6：実働時間は" & workHours
    '    Debug.Print "回帰テストNo.6：残業時間は" & overHours
    
    'イミディエイトウィンドウに貼り付けでテスト実行↓
    'Call CalcWorkTime(ThisWorkbook.Sheets("Attendance_2026_07"), 5)
    ' --- ここまでテスト用 ↑---
           
    ws.Cells(rowNum, ATT_COL_WORK_HOURS).value = workHours
    ws.Cells(rowNum, ATT_COL_OVERTIME).value = overHours

End Sub









'出勤日一覧作成
Public Function GetWorkdays(targetYear As Long, targetMonth As Long) As Collection
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim weekdays As New Collection
    Dim d As Variant
    Dim wsH As Worksheet
    Dim lr As Long
    Dim i As Long
    Dim workdays As New Collection

    '対象月の開始・終了
    monthStart = DateSerial(targetYear, targetMonth, 1)
    monthEnd = DateSerial(targetYear, targetMonth + 1, 0)
    
    '平日一覧を作成
    For d = monthStart To monthEnd
        If Weekday(d, vbMonday) <= 5 Then 'Weekdayの "vbMonday"は開始の曜日=1を表す。この場合「月」が1。金曜が「5」になる。<= 5つまり1~5は平日。
            weekdays.Add d
        End If
    Next d
    
    'シート存在チェック
    If Not SheetExists(SHEET_HOLIDAY) Then
        Set GetWorkdays = workdays   '空のコレクションを返す
        Exit Function
    End If
    
    Set wsH = ThisWorkbook.Sheets(SHEET_HOLIDAY)
    lr = GetLastRow(wsH, HM_COL_YEAR)
    
    '祝日一覧を作成（対象年のみ）
    Dim dictHolidays As Object
    Set dictHolidays = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lr
        If wsH.Cells(i, HM_COL_YEAR).value = targetYear Then
            dictHolidays(wsH.Cells(i, HM_COL_DATE).value) = True
        End If
    Next i

    '出勤日一覧を作成「平日一覧 - 祝日一覧」
    For Each d In weekdays
    
        If Not dictHolidays.Exists(d) Then
            workdays.Add d
        End If
        
    Next d

    Set GetWorkdays = workdays
   
End Function































