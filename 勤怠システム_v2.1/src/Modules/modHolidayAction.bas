Attribute VB_Name = "modHolidayAction"
Option Explicit

'==============================================
' modHolidayAction：休暇登録・取消・承認・却下
'==============================================

'休暇登録（有給・欠勤・特別休暇など）
Function RegisterLeave(empID As Long, STARTDate As Date, ENDDate As Date, leaveType As String, note As String) As result
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim lr As Long
    Dim StartDayRow As Long
    Dim EndDayRow As Long
    Dim d As Date
    Dim ApplicationDate As Date
    Dim y As Long, m As Long
    Dim Parts As Variant
    Dim found As Boolean
    Dim nameItem As Variant
    Dim GetName As Collection
    
    '初期値
    RegisterLeave = Other
    
    Set GetName = GetFutureAttendanceSheetsName
    ApplicationDate = Date
        
    For d = STARTDate To ENDDate
        found = False
        
        For Each nameItem In GetName
            Parts = Split(nameItem, "_")
            y = CLng(Parts(1))
            m = CLng(Parts(2))
            
            If year(d) = y And month(d) = m Then
            
                'シート存在チェック
                If Not SheetExists(nameItem) Then
                    RegisterLeave = Other 'ユーザーは一般含む
                    Exit For
                End If
            
                Set ws = ThisWorkbook.Sheets(nameItem)
                lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID) + 1
                
                ws.Cells(lr, ATT_COL_EMPLOYEE_ID).value = empID                ' 社員番号
                ws.Cells(lr, ATT_COL_DATE).value = d                           ' 日付
                ws.Cells(lr, ATT_COL_CLOCK_IN).value = ""                      ' 出勤
                ws.Cells(lr, ATT_COL_CLOCK_OUT).value = ""                     ' 退勤
                ws.Cells(lr, ATT_COL_LEAVE_TYPE).value = leaveType             ' 休暇種別
                ws.Cells(lr, ATT_COL_REMARKS).value = note                     ' 備考
                ws.Cells(lr, ATT_COL_APPLICATION_DATE).value = ApplicationDate ' 申請日
                ws.Cells(lr, ATT_COL_STATUS).value = "申請中"                  ' 申請状態
                
                found = True
                Exit For
            End If
        Next nameItem
        
        If Not found Then
            RegisterLeave = Other
            Exit Function
        End If
        
    Next d
    RegisterLeave = Success
    Call Sort_Attendance
    Exit Function

ErrHandler:
    RegisterLeave = Other '想定外の異常
End Function



'休暇申請取消処理
Public Function CancelLeaveRequest(selID As Long, selStaDate As Date, selEnDate As Date, selLT As String, selRemark As String) As result
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim lr As Long
    Dim d As Date
    Dim j As Long
    Dim sheetName As String
    Dim canCancel As Boolean
    
    '初期値
    CancelLeaveRequest = Other
    
    'まず全日が「申請中」かチェック
    canCancel = True

    For d = selStaDate To selEnDate
    
        sheetName = "Attendance_" & year(d) & "_" & Format(month(d), "00")
        
        'シート存在チェック
        If Not SheetExists(sheetName) Then
            CancelLeaveRequest = Other 'ユーザーは一般含む
            Exit Function
        End If
        
        Set ws = ThisWorkbook.Sheets(sheetName)
        Call ClearFilter(ws)
        lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)
    
            For j = 2 To lr
            
                If selID = ws.Cells(j, ATT_COL_EMPLOYEE_ID).value And d = CDate(ws.Cells(j, ATT_COL_DATE).value) _
                And selLT = ws.Cells(j, ATT_COL_LEAVE_TYPE).value And selRemark = ws.Cells(j, ATT_COL_REMARKS).value Then
                    '「申請中」でなければExit
                    If ws.Cells(j, ATT_COL_STATUS).value <> "申請中" Then
                        canCancel = False
                        Exit For
                    End If
                    
                End If
                
            Next j
        If Not canCancel Then Exit For

    Next d
    
    If Not canCancel Then
        CancelLeaveRequest = InvalidInput
        Exit Function
    End If
    
    '全日「申請中」なら削除実行
    For d = selStaDate To selEnDate

        sheetName = "Attendance_" & year(d) & "_" & Format(month(d), "00")
        Set ws = ThisWorkbook.Sheets(sheetName)
        Call ClearFilter(ws)
        lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)

        For j = lr To 2 Step -1
            If selID = ws.Cells(j, ATT_COL_EMPLOYEE_ID).value _
            And d = CDate(ws.Cells(j, ATT_COL_DATE).value) _
            And selLT = ws.Cells(j, ATT_COL_LEAVE_TYPE).value _
            And selRemark = ws.Cells(j, ATT_COL_REMARKS).value Then

                ws.Rows(j).Delete
                Exit For
            End If
        Next j
    Next d
       
    CancelLeaveRequest = Success
    Exit Function
    
ErrHandler:
    CancelLeaveRequest = Other '想定外の異常
End Function



'承認リクエスト
Public Function ApprovalRequest(selName As String, selStaDate As Date, selEnDate As Date, selType As String, selRemark As String) As result
    On Error GoTo ErrHandler
    
    Dim wsAtt As Worksheet
    Dim lrAtt As Long
    Dim d As Date
    Dim j As Long
    Dim TargetID As Long
    Dim sheetName As String

    '初期値
    ApprovalRequest = Other
    
    TargetID = GetEmployeeIDByName(selName) '辞書を用いた関数を使用
        
    For d = selStaDate To selEnDate
    
        sheetName = "Attendance_" & year(d) & "_" & Format(month(d), "00")
        
        'シート存在チェック
        If Not SheetExists(sheetName) Then
            ApprovalRequest = InvalidInput 'ユーザーは管理者Only
            Exit Function
        End If
        
        Set wsAtt = ThisWorkbook.Sheets(sheetName)
        lrAtt = GetLastRow(wsAtt, ATT_COL_EMPLOYEE_ID)

        For j = 2 To lrAtt
        
            If TargetID = wsAtt.Cells(j, ATT_COL_EMPLOYEE_ID).value And d = CDate(wsAtt.Cells(j, ATT_COL_DATE).value) _
                And selType = wsAtt.Cells(j, ATT_COL_LEAVE_TYPE).value And selRemark = wsAtt.Cells(j, ATT_COL_REMARKS).value Then
                
                wsAtt.Cells(j, ATT_COL_STATUS).value = "承認済"
                Exit For
            End If
        Next j
    Next d
    
    ApprovalRequest = Success
    Exit Function

ErrHandler:
    ApprovalRequest = Other '想定外の異常
End Function



'却下リクエスト
Public Function RejectionRequest(selName As String, selStaDate As Date, selEnDate As Date, selType As String, selRemark As String) As result
    On Error GoTo ErrHandler
    
    Dim wsAtt As Worksheet
    Dim lrAtt As Long
    Dim d As Date
    Dim j As Long
    Dim TargetID As Long
    Dim sheetName As String
    
    '初期値
    RejectionRequest = Other

    TargetID = GetEmployeeIDByName(selName) '辞書を用いた関数を使用


    For d = selStaDate To selEnDate

        sheetName = "Attendance_" & year(d) & "_" & Format(month(d), "00")
        
        'シート存在チェック
        If Not SheetExists(sheetName) Then
            RejectionRequest = InvalidInput 'ユーザーは管理者Only
            Exit Function
        End If
        
        Set wsAtt = ThisWorkbook.Sheets(sheetName)
        lrAtt = GetLastRow(wsAtt, ATT_COL_EMPLOYEE_ID)
            
        For j = 2 To lrAtt
            If TargetID = wsAtt.Cells(j, ATT_COL_EMPLOYEE_ID).value And d = CDate(wsAtt.Cells(j, ATT_COL_DATE).value) _
            And selType = wsAtt.Cells(j, ATT_COL_LEAVE_TYPE).value And selRemark = wsAtt.Cells(j, ATT_COL_REMARKS).value Then

                wsAtt.Cells(j, ATT_COL_STATUS).value = "却下済"
                Exit For
            End If
        Next j
    Next d

    RejectionRequest = Success
    Exit Function

ErrHandler:
    RejectionRequest = Other '想定外の異常
End Function


