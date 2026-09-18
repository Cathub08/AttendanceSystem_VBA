Attribute VB_Name = "modAttendanceReport"
Option Explicit

'==================================
' modAttendanceReport：勤怠レポート
'==================================


'管理者用：今月分の全社員の勤怠を再計算
Public Function RecalcAllAttendance() As result
    On Error GoTo ErrHandler
    
'    'TEST↓--------------------------------------------
'    Debug.Print "【No6-2】RecalcAllAttendance に到達"
'    'TEST↑--------------------------------------------
    
    '初期値
    RecalcAllAttendance = Other
    
    Dim sheetName As String
    sheetName = GetAttendanceSheetName

    'シート存在チェック
    If Not SheetExists(sheetName) Then '想定内の異常
        RecalcAllAttendance = InvalidInput
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim lr As Long
    lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)

    Dim i As Long
    Dim updated As Boolean
    updated = False
    
    For i = 2 To lr
        If ws.Cells(i, ATT_COL_CLOCK_IN).value <> "" And ws.Cells(i, ATT_COL_CLOCK_OUT).value <> "" Then
            Call CalcWorkTime(ws, i)
            updated = True
        End If
    Next i
        
    If updated Then
        RecalcAllAttendance = Success
    Else
        RecalcAllAttendance = InvalidInput '今月の記録がない状態でのボタン押下
    End If
    
    Exit Function
    
ErrHandler:
    RecalcAllAttendance = Other '想定外の異常
End Function


'未来分を含めてattendanceシートをソートする
Public Sub Sort_Attendance()
    Dim lr As Long
    Dim ws As Worksheet
    Dim nameItem As Variant
    Dim GetName As Collection
    Set GetName = GetFutureAttendanceSheetsName
                       
    For Each nameItem In GetName
    
        'シート存在チェック
        If Not SheetExists(nameItem) Then
            '存在しないシートはスキップ
            GoTo ContinueLoop
        End If
        
        Set ws = ThisWorkbook.Sheets(nameItem)
        lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)

        With ws.Sort
           .SortFields.Clear
           .SortFields.Add key:=ws.Cells(2, ATT_COL_DATE), SortOn:=xlSortOnValues, Order:=xlAscending
           .SortFields.Add key:=ws.Cells(2, ATT_COL_EMPLOYEE_ID), SortOn:=xlSortOnValues, Order:=xlAscending
           .SetRange ws.Range(ws.Cells(2, ATT_COL_EMPLOYEE_ID), ws.Cells(lr, ATT_COL_STATUS))
           .Header = xlNo
           .Apply
        End With
    
ContinueLoop:
    Next nameItem

End Sub


' ログイン中社員の当月勤怠をフィルタ表示
Public Function ShowAttendanceList() As result
    On Error GoTo ErrHandler
    
    '初期値
    ShowAttendanceList = Other
    
    'シート存在チェック
    If Not SheetExists(GetAttendanceSheetName) Then
        ShowAttendanceList = Other
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(GetAttendanceSheetName)

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    Dim lr As Long
    lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)

    ws.Range("A1:K" & lr).AutoFilter Field:=1, Criteria1:=LoggedInID
    ws.Range("A1:K" & lr).AutoFilter Field:=2, _
        Criteria1:=">=" & DateSerial(year(Date), month(Date), 1), _
        Operator:=xlAnd, _
        Criteria2:="<=" & DateSerial(year(Date), month(Date) + 1, 0)

    ws.Activate
    ShowAttendanceList = Success
    Exit Function

ErrHandler:
    ShowAttendanceList = Other '想定外の異常
End Function


'===============================
' 月次集計レポート（Excel シート作成）
'===============================
Public Function GenerateMonthlyReport() As result
    On Error GoTo ErrHandler
    
    Dim wsA As Worksheet
    Dim wsR As Worksheet
    Dim lr As Long
    Dim reportRow As Long
    Dim empID As Variant
    Dim empName As String
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim Dict As Object
    Dim key As Variant
    Dim ThisMon As String
    Dim Titl As String 'タイトルを格納
    Dim reportName As String
    Dim dictEmp As Object
    
    '初期値
    GenerateMonthlyReport = Other

    '今月の開始・終了日
    monthStart = DateSerial(year(Date), month(Date), 1)
    monthEnd = DateSerial(year(Date), month(Date) + 1, 0)

    reportName = "MonthlyReport_" & Format(Date, "yyyymm")
    
    '既存の同名レポートシートがあれば削除
    If SheetExists(reportName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(reportName).Delete
        Application.DisplayAlerts = True
    End If

'TEST【ModRepo:No9,MainMenu:No8】↓-------------
'Stop 'ProjectExplorerでシート削除を確認
'TEST↑-----------------------------------------

    'シート存在チェック
    If Not SheetExists(GetAttendanceSheetName) Then
        GenerateMonthlyReport = Other
        Exit Function
    End If
    
    Set wsA = ThisWorkbook.Sheets(GetAttendanceSheetName)

    'シート存在チェック
    If Not SheetExists("MonthlyReport_Template") Then
        GenerateMonthlyReport = Other
        Exit Function
    End If

    lr = GetLastRow(wsA, ATT_COL_EMPLOYEE_ID)
    
    If lr = 1 Then
        MsgBox "今月の勤怠シートの内容がありません。月次レポートは出力されません。", vbCritical
        GenerateMonthlyReport = NoData
        Exit Function
    End If

    'テンプレをコピーし新しいレポートシート作成
    ThisWorkbook.Sheets("MonthlyReport_Template").Copy _
    After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    Set wsR = ActiveSheet
    wsR.name = reportName

    wsR.Cells.Font.name = "Meiryo UI"
    wsR.Cells.Font.Size = 10

    ThisMon = Format(Date, "yyyy年mm月")
    Titl = "勤怠レポート（" & ThisMon & "度）"
    wsR.Range("A1").value = Titl
    wsR.Range("A1:H1").Merge
    wsR.Range("A1").HorizontalAlignment = xlCenter
    wsR.Range("A1").Font.Size = 14
    wsR.Range("A1").Font.Bold = True
    wsR.Range("A1").Interior.Color = RGB(44, 62, 80)
    wsR.Range("A1").Font.Color = RGB(255, 255, 255)
    wsR.Rows(2).Font.Size = 11
    
    'ヘッダー作成
    wsR.Range("A2").value = "社員ID"
    wsR.Range("B2").value = "氏名"
    wsR.Range("C2").value = "出勤日数"
    wsR.Range("D2").value = "欠勤日数"
    wsR.Range("E2").value = "有給取得日数"
    wsR.Range("F2").value = "勤務時間合計"
    wsR.Range("G2").value = "残業時間合計"
    wsR.Range("H2").value = "総労働時間"

    wsR.Range("A2:H2").Font.Bold = True
    wsR.Range("A2:H2").Interior.Color = RGB(44, 62, 80)
    wsR.Range("A2:H2").Font.Color = RGB(255, 255, 255)

    Dim y As Long, m As Long
    Dim d As Variant
    Dim workdaysList As Collection
    Dim i As Long
    Dim j As Long
    Dim arr '型宣言省略(=Variant)

    y = year(Date)
    m = month(Date)

    Set workdaysList = GetWorkdays(y, m)
    
    ' --- 統合テスト用コード ↓---
'    Debug.Print "出勤想定日数は" & workdaysList.Count & "日" 'Test No3確認
'    Stop
    ' --- ここまでテスト用 ↑---

  
    '辞書で集計
    Set Dict = CreateObject("Scripting.Dictionary")
    
    Set dictEmp = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lr
        empID = wsA.Cells(i, ATT_COL_EMPLOYEE_ID).value
        If Not dictEmp.Exists(empID) Then
            dictEmp.Add empID, True '重複なしで今月のATTシートのempIDをdictEmpのKeyに追加する
        End If
    Next i
        
    For Each empID In dictEmp.Keys
              
        '辞書にキーがなければ初期化(Key=empIDとItem=配列[要素数8個、要素の中身0]をdictに追加する)
        If Not Dict.Exists(empID) Then
            Dict.Add empID, Array(0, 0, 0, 0, 0, 0, 0, 0) '出勤日数, 欠勤日数,有給取得日数,勤務時間合計,残業時間合計, 総労働時間,特別休暇,代休
        End If
             
        arr = Dict(empID)

        'Attendance シートを検索
        For i = 2 To lr
    
            '対象月 & 対象社員
            If wsA.Cells(i, ATT_COL_EMPLOYEE_ID).value = empID And _
               wsA.Cells(i, ATT_COL_DATE).value >= monthStart And _
               wsA.Cells(i, ATT_COL_DATE).value <= monthEnd Then
                
                '出勤日数（休日出勤も含む）
                If wsA.Cells(i, ATT_COL_CLOCK_IN).value <> "" Then
                    arr(0) = arr(0) + 1
                End If
    
                '"有給取得日数"
                If wsA.Cells(i, ATT_COL_LEAVE_TYPE).value = "有給休暇" Then
                    arr(2) = arr(2) + 1
                    arr(5) = arr(5) + STD_WORK_HOURS
                End If
    
                 '特別休暇
                If wsA.Cells(i, ATT_COL_LEAVE_TYPE).value = "特別休暇" Then
                    arr(6) = arr(6) + 1 'arr(6)は現在使用しないが、後ほど必要になったら処理を追加可能。
                End If
                
                '代休
                If wsA.Cells(i, ATT_COL_LEAVE_TYPE).value = "代休" Then
                    arr(7) = arr(7) + 1 'arr(7)は現在使用しないが、後ほど必要になったら処理を追加可能。
                End If
                                                                                       
                '勤務時間（有給・特別休暇・代休以外）
                If wsA.Cells(i, ATT_COL_LEAVE_TYPE).value <> "有給休暇" _
                And wsA.Cells(i, ATT_COL_LEAVE_TYPE).value <> "特別休暇" _
                And wsA.Cells(i, ATT_COL_LEAVE_TYPE).value <> "代休" Then
                
                    arr(3) = arr(3) + wsA.Cells(i, ATT_COL_WORK_HOURS).value '勤務時間合計
                    arr(4) = arr(4) + wsA.Cells(i, ATT_COL_OVERTIME).value '残業時間合計
                    arr(5) = arr(5) + wsA.Cells(i, ATT_COL_WORK_HOURS).value + wsA.Cells(i, ATT_COL_OVERTIME).value '総労働時間
                End If
                                                          
            End If
    
        Next i
        
          
        '欠勤を計算し直す（Workdays - (出勤 + 有給 + 特別休暇+代休)）
        Dim calcAbs As Long
        calcAbs = workdaysList.Count - (arr(0) + arr(2) + arr(6) + arr(7))
        
        If calcAbs < 0 Then calcAbs = 0
        
        arr(1) = calcAbs
                                   
        Dict(empID) = arr
    
    Next empID
       
    '辞書を配列にコピーする
    Dim AttID() As Variant
    Dim idx As Long
    
    ReDim AttID(0 To Dict.Count - 1) '配列サイズを指定
    
    idx = 0
    For Each key In Dict.Keys
        AttID(idx) = key
        idx = idx + 1 'idxは配列のインデックス(配列番号)。次の要素を入れるのに次の引き出し番号を指定する必要がある。
    Next key
    
    '配列を昇順に入れ替える
    Dim tempID As Variant
    
    For i = LBound(AttID) To UBound(AttID) - 1
        For j = i + 1 To UBound(AttID)
            If AttID(i) > AttID(j) Then
                tempID = AttID(i)
                AttID(i) = AttID(j)
                AttID(j) = tempID
            End If
        Next j
    Next i
       
           
    Dim eID As Long
    'レポート書き込み
    reportRow = 3

    '配列を使用してレポート出力
    For i = LBound(AttID) To UBound(AttID)

        eID = AttID(i)
        empName = GetEmployeeNameByID(eID) '社員名取得

        If empName = "" Then
            empName = "社員名未登録"
        End If

        wsR.Cells(reportRow, MRepo_COL_EMPLOYEE_ID).value = eID                    '"社員ID"
        wsR.Cells(reportRow, MRepo_COL_EMPLOYEE_NAME).value = empName              '"氏名"
        wsR.Cells(reportRow, MRepo_COL_NUMBER_OF_WORKING_DAY).value = Dict(eID)(0) '"出勤日数"
        wsR.Cells(reportRow, MRepo_COL_NUMBER_OF_DAYS_ABSENT).value = Dict(eID)(1) '"欠勤日数"
        wsR.Cells(reportRow, MRepo_COL_PAID_LEAVE_DAYS).value = Dict(eID)(2)       '"有給取得日数"
        wsR.Cells(reportRow, MRepo_COL_TOTAL_WORKING_HOURS).value = Dict(eID)(3)   '"勤務時間合計"
        wsR.Cells(reportRow, MRepo_COL_TOTAL_OVERTIME_HOURS).value = Dict(eID)(4)  '"残業時間合計"
        wsR.Cells(reportRow, MRepo_COL_TOTAL_HOURS).value = Dict(eID)(5)           '"総労働時間"

        reportRow = reportRow + 1

    Next i
    
    '列幅調整
    wsR.Columns("A:H").AutoFit
    wsR.Activate
    GenerateMonthlyReport = Success
    Exit Function

ErrHandler:
    GenerateMonthlyReport = Other '想定外の異常
End Function


