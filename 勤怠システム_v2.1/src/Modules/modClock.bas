Attribute VB_Name = "modClock"
Option Explicit

'================
' modClock：打刻
'================

'今日の出勤記録を取得
Public Function GetClockIn(empID As Long, today As Date) As Variant
    Dim lr As Long
    Dim i As Long
    Dim ws As Worksheet
    
    'シート存在チェック
    If Not SheetExists(GetAttendanceSheetName) Then Exit Function
    
    Set ws = ThisWorkbook.Sheets(GetAttendanceSheetName)
    lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)
    
    'フィルタ解除
    Call ClearFilter(ws)
    
    For i = lr To 2 Step -1
        If ws.Cells(i, ATT_COL_EMPLOYEE_ID).value = empID And ws.Cells(i, ATT_COL_DATE).value = today Then
            If ws.Cells(i, ATT_COL_CLOCK_IN).value <> "" Then
                GetClockIn = ws.Cells(i, ATT_COL_CLOCK_IN).value
                Exit Function
            Else
                GetClockIn = ""
            End If
        End If
    Next i
    
End Function

'今日の退勤記録を取得
Public Function GetClockOut(empID As Long, today As Date) As Variant
    Dim lr As Long
    Dim i As Long
    Dim ws As Worksheet
    
    'シート存在チェック
    If Not SheetExists(GetAttendanceSheetName) Then Exit Function
    
    Set ws = ThisWorkbook.Sheets(GetAttendanceSheetName)
    lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)
    
    'フィルタ解除
    Call ClearFilter(ws)
    
    For i = lr To 2 Step -1
        If ws.Cells(i, ATT_COL_EMPLOYEE_ID).value = empID And ws.Cells(i, ATT_COL_DATE).value = today Then
            If ws.Cells(i, ATT_COL_CLOCK_OUT).value <> "" Then
                GetClockOut = ws.Cells(i, ATT_COL_CLOCK_OUT).value
                Exit Function
            Else
                GetClockOut = ""
            End If
        End If
    Next i
    
End Function

'今日の休憩記録を取得
Public Function GetBreakTime(empID As Long, today As Date) As Variant
    Dim lr As Long
    Dim i As Long
    Dim ws As Worksheet
    
    'シート存在チェック
    If Not SheetExists(GetAttendanceSheetName) Then Exit Function
    
    Set ws = ThisWorkbook.Sheets(GetAttendanceSheetName)
    lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)
    
    'フィルタ解除
    Call ClearFilter(ws)
    
    For i = lr To 2 Step -1
        If ws.Cells(i, ATT_COL_EMPLOYEE_ID).value = empID And ws.Cells(i, ATT_COL_DATE).value = today Then
            If ws.Cells(i, ATT_COL_BREAK_TIME).value <> "" Then
                GetBreakTime = ws.Cells(i, ATT_COL_BREAK_TIME).value
                Exit Function
            Else
                GetBreakTime = ""
            End If
        End If
    Next i
End Function

'出勤打刻を記録
Public Function ClockIn(empID As Long, today As Date) As result
    On Error GoTo ErrHandler
    Dim lr As Long
    Dim i As Long
    Dim ws As Worksheet
    
    '初期値
    ClockIn = Other
    
    'シート存在チェック
    If Not SheetExists(GetAttendanceSheetName) Then Exit Function
    
    Set ws = ThisWorkbook.Sheets(GetAttendanceSheetName)
    lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)
    
    'フィルタ解除
    Call ClearFilter(ws)
    
    For i = lr To 2 Step -1
        If ws.Cells(i, ATT_COL_EMPLOYEE_ID).value = empID And ws.Cells(i, ATT_COL_DATE).value = today Then
            If ws.Cells(i, ATT_COL_CLOCK_IN).value <> "" Then
                ClockIn = InvalidInput
                Exit Function
            End If
        End If
    Next i

    lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID) + 1
    ws.Cells(lr, ATT_COL_EMPLOYEE_ID).value = empID
    ws.Cells(lr, ATT_COL_DATE).value = today
    ws.Cells(lr, ATT_COL_CLOCK_IN).value = Format(Now, "h:mm:ss")
    
    ClockIn = Success
    Exit Function

ErrHandler:
    ClockIn = Other '想定外の異常
End Function

'退勤打刻を記録
Public Function ClockOut(empID As Long, today As Date) As result
    On Error GoTo ErrHandler
    Dim lr As Long
    Dim i As Long
    Dim ws As Worksheet
    
    '初期値
    ClockOut = Other
    
    'シート存在チェック
    If Not SheetExists(GetAttendanceSheetName) Then Exit Function
    
    Set ws = ThisWorkbook.Sheets(GetAttendanceSheetName)
    lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)

    'フィルタ解除
    Call ClearFilter(ws)

    For i = lr To 2 Step -1
        If ws.Cells(i, ATT_COL_EMPLOYEE_ID).value = empID And ws.Cells(i, ATT_COL_DATE).value = today Then
            If ws.Cells(i, ATT_COL_CLOCK_OUT).value <> "" Then
                ClockOut = InvalidInput
                Exit Function
            End If

            If ws.Cells(i, ATT_COL_CLOCK_IN).value = "" Then
                ClockOut = NoRecordsForToday
                Exit Function
            End If
            
            ws.Cells(i, ATT_COL_CLOCK_OUT).value = Format(Now, "h:mm:ss")
            Call CalcWorkTime(ws, i)
            ClockOut = Success
            Exit Function
        End If
    Next i

ErrHandler:
    ClockOut = Other '想定外の異常
End Function


'休憩時間を記録
Public Function ReflectBreakTimes(empID As Long, today As Date, BREAK_TIME As Date) As result
    On Error GoTo ErrHandler
    Dim lr As Long
    Dim i As Long
    Dim ws As Worksheet
      
    '初期値
    ReflectBreakTimes = Other
      
    'シート存在チェック
    If Not SheetExists(GetAttendanceSheetName) Then Exit Function
      
    Set ws = ThisWorkbook.Sheets(GetAttendanceSheetName)
    lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)

    'フィルタ解除
    Call ClearFilter(ws)

    For i = lr To 2 Step -1
        If ws.Cells(i, ATT_COL_EMPLOYEE_ID).value = empID And ws.Cells(i, ATT_COL_DATE).value = today Then
            ws.Cells(i, ATT_COL_BREAK_TIME).value = BREAK_TIME
                If ws.Cells(i, ATT_COL_CLOCK_OUT).value <> "" Then
                    Call CalcWorkTime(ws, i)
                End If
                ReflectBreakTimes = Success
                Exit Function
        End If
    Next i
            
    ReflectBreakTimes = NoRecordsForToday '本日の出勤記録無し
    Exit Function
            
ErrHandler:
    ReflectBreakTimes = Other '想定外の異常
End Function



