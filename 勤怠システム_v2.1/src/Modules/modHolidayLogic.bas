Attribute VB_Name = "modHolidayLogic"
Option Explicit

'=====================================
' modHolidayLogic：休暇検索＋連休判定
'=====================================

'休暇の申請を探す
Public Sub SearchHolidayRequests( _
ByVal selName As String, _
ByVal selID As Long, _
ByVal selStatus As String, _
ByVal selStaDate As Date, _
ByVal selEnDate As Date, _
ByRef stDate() As Date, _
ByRef enDate() As Date, _
ByRef TargetID() As Long, _
ByRef Types() As String, _
ByRef Remarks() As String)

    Dim wsAtt As Worksheet
    Dim lrAtt As Long
    Dim STARTDate As Date, ENDDate As Date
    Dim i As Long
    Dim j As Long
    Dim Count As Long
    Dim Cnt As Long
    Dim nameItem As Variant
    Dim dates() As Date
    Dim lvType() As String
    Dim RM() As String
    Dim eID() As Long
    Dim holidays As Collection
    Dim leaveType As Collection
    Dim Remark As Collection
    Dim empID As Collection
    Dim N As Variant
    Dim pass As Boolean
    Dim d As Date
    
    Set holidays = New Collection
    Set leaveType = New Collection
    Set Remark = New Collection
    Set empID = New Collection
    
    i = 0
    j = 0
    Count = 0
             
    For Each nameItem In TgName 'Thisworkbook全量のAttendanceシート
       
    Set wsAtt = ThisWorkbook.Sheets(nameItem)
    lrAtt = GetLastRow(wsAtt, ATT_COL_EMPLOYEE_ID)

        For i = 2 To lrAtt '行
            pass = True
            If wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value <> "" Then
                
                '① 社員名フィルタ
                If selID <> 0 Then
                    If wsAtt.Cells(i, ATT_COL_EMPLOYEE_ID).value <> selID Then
                        pass = False
                    End If
                End If
                
                '② ステータスフィルタ
                If selStatus <> "" Then
                    If wsAtt.Cells(i, ATT_COL_STATUS).value <> selStatus Then
                        pass = False
                    End If
                End If
                
                '③ 日付フィルタ
                If selStaDate <> 0 And selEnDate <> 0 Then
                    d = wsAtt.Cells(i, ATT_COL_DATE).value
                    If d < selStaDate Or d > selEnDate Then
                        pass = False
                    End If
                End If
                
                'すべての条件を満たしたら Add
                If pass Then
                    holidays.Add CDate(wsAtt.Cells(i, ATT_COL_DATE).value)  'ATTシートの日付
                    leaveType.Add wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value  'ATTシートの休暇種別
                    Remark.Add wsAtt.Cells(i, ATT_COL_REMARKS).value        'ATTシートの備考
                    empID.Add wsAtt.Cells(i, ATT_COL_EMPLOYEE_ID).value     'ATTシートの社員番号
                End If
                        
            End If

        Next i
    Next nameItem

    If holidays.Count = 0 Then
        Exit Sub
    End If
 
    'コレクションを配列へ入れ替え
    If holidays.Count > 0 Then
        ReDim dates(1 To holidays.Count)
        ReDim lvType(1 To leaveType.Count)
        ReDim RM(1 To Remark.Count)
        ReDim eID(1 To empID.Count)
        i = 1

            For Each N In holidays    'ATTシートの日付
                dates(i) = N
                i = i + 1
            Next N

            i = 1
            For Each N In leaveType   'ATTシートの休暇種別
                lvType(i) = N
                i = i + 1
            Next N
            
            i = 1
            For Each N In Remark   'ATTシートの備考
                RM(i) = N
                i = i + 1
            Next N

            i = 1
            For Each N In empID       'ATTシートの社員番号
                eID(i) = N
                i = i + 1
            Next N

    End If
        
   'バブルソート[ID昇順、同じなら日付昇順]
    Call SortFourArrays(eID, dates, lvType, RM)
       
' --- 連休判定 ---
    ReDim stDate(1 To 1)
    ReDim enDate(1 To 1)
    ReDim TargetID(1 To 1)
    ReDim Types(1 To 1)
    ReDim Remarks(1 To 1)
    Cnt = 1
    
    STARTDate = dates(1)    '一つ目の要素を格納
    ENDDate = dates(1)      '一つ目の要素を格納
    TargetID(Cnt) = eID(1)  '一つ目の要素を格納
    Types(Cnt) = lvType(1)  '一つ目の要素を格納
    Remarks(Cnt) = RM(1)  '一つ目の要素を格納

    For j = 2 To UBound(dates)
        
        If eID(j) = eID(j - 1) And lvType(j) = lvType(j - 1) And RM(j) = RM(j - 1) And dates(j) - ENDDate = 1 Then
            '連続している場合
            ENDDate = dates(j)
        Else
            '連続していない場合' 連休終了 → 出力
            stDate(Cnt) = STARTDate
            enDate(Cnt) = ENDDate
            TargetID(Cnt) = eID(j - 1)
            Types(Cnt) = lvType(j - 1)
            Remarks(Cnt) = RM(j - 1)
            
            Cnt = Cnt + 1
            ReDim Preserve stDate(1 To Cnt)
            ReDim Preserve enDate(1 To Cnt)
            ReDim Preserve TargetID(1 To Cnt)
            ReDim Preserve Types(1 To Cnt)
            ReDim Preserve Remarks(1 To Cnt)
            
            '新しい連休開始
            STARTDate = dates(j)
            ENDDate = dates(j)
        End If
    Next j

    '最後の連休を出力
    stDate(Cnt) = STARTDate
    enDate(Cnt) = ENDDate
    TargetID(Cnt) = eID(UBound(eID))
    Types(Cnt) = lvType(UBound(lvType))
    Remarks(Cnt) = RM(UBound(RM))
    
End Sub



'休暇開始日・終了日を取得する
Public Sub GetConsecutiveHolidays(ByRef stDate() As Date, ByRef enDate() As Date)
    Dim ws As Worksheet
    Dim lr As Long
    Dim STARTDate As Date, ENDDate As Date
    Dim i As Long
    Dim j As Long
    Dim Count As Long
    Dim NotNull() As Date
    Dim Cnt As Long
    Dim LT() As String
    Dim nameItem As Variant
    Dim GetName As Collection
    Dim ConsecutiveHolidays As Collection
    Dim leaveType As Collection
    Dim N As Variant
    Dim temp As Variant
    Dim empID As Long
    Dim Remark As Collection
    Dim RM() As String
    
    Set ConsecutiveHolidays = New Collection
    Set leaveType = New Collection
    Set Remark = New Collection
    Set GetName = GetFutureAttendanceSheetsName
    empID = GetLoggedInID()

    i = 0
    j = 0
    Count = 0
        
    For Each nameItem In GetName
        
        If SheetExists(nameItem) Then 'シート存在チェック
    
            Set ws = ThisWorkbook.Sheets(nameItem)
            Call ClearFilter(ws)
            lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)
        
            For i = 2 To lr ' 1行目は見出し
                    If empID = ws.Cells(i, ATT_COL_EMPLOYEE_ID).value And Date < CDate(ws.Cells(i, ATT_COL_DATE).value) _
                    And ws.Cells(i, ATT_COL_LEAVE_TYPE).value <> "" Then
                        ConsecutiveHolidays.Add CDate(ws.Cells(i, ATT_COL_DATE).value)
                        leaveType.Add ws.Cells(i, ATT_COL_LEAVE_TYPE).value
                        Remark.Add ws.Cells(i, ATT_COL_REMARKS).value
                    End If
            Next i
        End If
    Next nameItem


    If ConsecutiveHolidays.Count > 0 Then
    
        ReDim NotNull(1 To ConsecutiveHolidays.Count)
        ReDim LT(1 To ConsecutiveHolidays.Count)
        ReDim RM(1 To ConsecutiveHolidays.Count)

        i = 1

            For Each N In ConsecutiveHolidays
                NotNull(i) = N
                i = i + 1
            Next N
            
            N = 0
            i = 1
            For Each N In leaveType
                LT(i) = N
                i = i + 1
            Next N
            
            N = 0
            i = 1
            For Each N In Remark
                RM(i) = N
                i = i + 1
            Next N

    Else
        Exit Sub
    End If
    
    '--- バブルソートで昇順に並べ替え ---
    For i = LBound(NotNull) To UBound(NotNull) - 1
        For j = i + 1 To UBound(NotNull)

            If NotNull(i) > NotNull(j) Then
                '--- 日付の入れ替え ---
                temp = NotNull(i)
                NotNull(i) = NotNull(j)
                NotNull(j) = temp
                
                '--- 種別の入れ替え---
                temp = LT(i)
                LT(i) = LT(j)
                LT(j) = temp
                
                '--- 備考の入れ替え---
                temp = RM(i)
                RM(i) = RM(j)
                RM(j) = temp
                
            End If
        Next j
    Next i
    
' --- 連休判定 ---
    ReDim stDate(1 To 1)
    ReDim enDate(1 To 1)
    Cnt = 1
    
    STARTDate = NotNull(1) '一つ目の要素を格納
    ENDDate = NotNull(1) '一つ目の要素を格納

    For j = 2 To UBound(NotNull)

        If LT(j) = LT(j - 1) And RM(j) = RM(j - 1) And NotNull(j) - ENDDate = 1 Then
            '連続している場合
            ENDDate = NotNull(j)
        Else
            '連続していない場合' 連休終了 → 出力
            stDate(Cnt) = STARTDate
            enDate(Cnt) = ENDDate
            
            Cnt = Cnt + 1
            ReDim Preserve stDate(1 To Cnt)
            ReDim Preserve enDate(1 To Cnt)
            
            '新しい連休開始
            STARTDate = NotNull(j)
            ENDDate = NotNull(j)
        End If
    Next j

    '最後の連休を出力
    stDate(Cnt) = STARTDate
    enDate(Cnt) = ENDDate
    
End Sub



'全ユーザーの休暇開始日・終了日を取得
Public Sub GetConsecutiveHolidaysOfAllUsers(ByRef stDate() As Date, ByRef enDate() As Date, ByRef TargetID() As Long)
    Dim wsAtt As Worksheet
    Dim lrAtt As Long
    Dim STARTDate As Date, ENDDate As Date
    Dim i As Long
    Dim j As Long
    Dim Cnt As Long
    Dim dates() As Date
    Dim Types() As String
    Dim RM() As String
    Dim eID() As Long
    Dim nameItem As Variant
    Dim GetName As Collection
    Dim idx As Long

    '初期化
    i = 0
    j = 0
    Cnt = 0

    Set GetName = GetFutureAttendanceSheetsName
    For Each nameItem In GetName
    
        If SheetExists(nameItem) Then 'シート存在チェック
            Set wsAtt = ThisWorkbook.Sheets(nameItem)
            
            Call ClearFilter(wsAtt)
            
            lrAtt = GetLastRow(wsAtt, ATT_COL_EMPLOYEE_ID)
            For i = 2 To lrAtt ' 1行目は見出し
                              
                If wsAtt.Cells(i, ATT_COL_STATUS).value = "申請中" _
                And Date < CDate(wsAtt.Cells(i, ATT_COL_DATE).value) _
                And wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value <> "" Then
                Cnt = Cnt + 1

                End If
                
            Next i
        End If
    Next nameItem
           
    '申請0件ならExit
    If Cnt = 0 Then
        Exit Sub
    End If
           
    '配列のサイズを確定
    ReDim dates(1 To Cnt)
    ReDim Types(1 To Cnt)
    ReDim RM(1 To Cnt)
    ReDim eID(1 To Cnt)
       
    idx = 1
    '未来分含めたシートをセット
    For Each nameItem In GetName
    
        Set wsAtt = ThisWorkbook.Sheets(nameItem)
        lrAtt = GetLastRow(wsAtt, ATT_COL_EMPLOYEE_ID)
    
        For i = 2 To lrAtt
            If wsAtt.Cells(i, ATT_COL_STATUS).value = "申請中" _
            And CDate(wsAtt.Cells(i, ATT_COL_DATE).value) > Date _
                            And wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value <> "" Then
                dates(idx) = CDate(wsAtt.Cells(i, ATT_COL_DATE).value)
                Types(idx) = wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value
                RM(idx) = wsAtt.Cells(i, ATT_COL_REMARKS).value
                eID(idx) = wsAtt.Cells(i, ATT_COL_EMPLOYEE_ID).value
                idx = idx + 1
            End If
        Next i
    
    Next nameItem

    Call SortFourArrays(eID, dates, Types, RM)
     
    '連休判定
    ReDim stDate(1 To 1)
    ReDim enDate(1 To 1)
    ReDim TargetID(1 To 1)
    Cnt = 1
    
    STARTDate = dates(1)   '一つ目の要素を格納
    ENDDate = dates(1)     '一つ目の要素を格納
    TargetID(Cnt) = eID(1) '一つ目の要素を格納

    For j = 2 To UBound(dates)
        
        If eID(j) = eID(j - 1) And Types(j) = Types(j - 1) And RM(j) = RM(j - 1) And dates(j) - ENDDate = 1 Then
            '連続している場合
            ENDDate = dates(j)
        Else
            '連続していない場合' 連休終了 → 出力
            stDate(Cnt) = STARTDate
            enDate(Cnt) = ENDDate
            TargetID(Cnt) = eID(j - 1)
            
            Cnt = Cnt + 1
            ReDim Preserve stDate(1 To Cnt)
            ReDim Preserve enDate(1 To Cnt)
            ReDim Preserve TargetID(1 To Cnt)
            
            '新しい連休開始
            STARTDate = dates(j)
            ENDDate = dates(j)
        End If
    Next j

    '最後の連休を出力
    stDate(Cnt) = STARTDate
    enDate(Cnt) = ENDDate
    TargetID(Cnt) = eID(UBound(eID))
    
End Sub



'全ユーザー・全ステータスの休暇開始日・終了日を取得する
Public Sub GetHolidayFromAllUsersAndAllStatuses _
(ByRef stDate() As Date, ByRef enDate() As Date, ByRef TargetID() As Long, ByRef Types() As String, ByRef Remarks() As String)
    Dim wsAtt As Worksheet
    Dim lrAtt As Long
    Dim nameItem As Variant
    Dim Dict As Object
    Dim empID As Variant
    Dim d As Date
    Dim lvType As String
    Dim Remark As String
    Dim empDates As Object
    Dim empTypes As Object
    Dim empRemark As Object
    Dim i As Long
    Dim j As Long
    Dim col As Collection
    Dim HolidaysArr() As Variant
    Dim LTArr() As Variant
    Dim RmArr() As Variant
    Dim empIDArr() As Long
    Dim datesArr() As Date
    Dim typesArr() As String
    Dim remarksArr() As String
    Dim outCnt As Long
    Dim curID As Long
    Dim curStart As Date
    Dim curEnd As Date
    Dim curType As String
    Dim curRemarks As String
    
    Set col = New Collection

    '辞書：EmpID → 日付配列
    Set Dict = CreateObject("Scripting.Dictionary")
        
    '配列のサイズを測る
    For Each nameItem In TgName 'Thisworkbook全量のAttendanceSheets名
    
        If SheetExists(nameItem) Then 'シート存在チェック
            Set wsAtt = ThisWorkbook.Sheets(nameItem)
            Call ClearFilter(wsAtt)
            lrAtt = GetLastRow(wsAtt, ATT_COL_EMPLOYEE_ID)
            
            For i = 2 To lrAtt '1行目は見出し
                If wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value <> "" Then
                
                    empID = wsAtt.Cells(i, ATT_COL_EMPLOYEE_ID).value
                    d = CDate(wsAtt.Cells(i, ATT_COL_DATE).value)
                    lvType = wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value
                    Remark = wsAtt.Cells(i, ATT_COL_REMARKS).value

                    '辞書にキーがなければ作成
                     If Not Dict.Exists(empID) Then
                         Set empDates = CreateObject("Scripting.Dictionary")
                         Set empTypes = CreateObject("Scripting.Dictionary")
                         Set empRemark = CreateObject("Scripting.Dictionary")
                         Dict.Add empID, Array(empDates, empTypes, empRemark)
                     End If
    
                     '日付と種別を追加
                     Dict(empID)(0)(Dict(empID)(0).Count) = d
                     Dict(empID)(1)(Dict(empID)(1).Count) = lvType
                     Dict(empID)(2)(Dict(empID)(2).Count) = Remark

                End If
            Next i
        End If
    Next nameItem
        
    For Each empID In Dict.Keys
        HolidaysArr = Dict(empID)(0).Items
        LTArr = Dict(empID)(1).Items
        RmArr = Dict(empID)(2).Items
        
        For i = 0 To Dict(empID)(0).Count - 1
            col.Add Array(empID, HolidaysArr(i), LTArr(i), RmArr(i))
        Next i
    Next empID
       
    ReDim empIDArr(1 To col.Count)
    ReDim datesArr(1 To col.Count)
    ReDim typesArr(1 To col.Count)
    ReDim remarksArr(1 To col.Count)
    
    For i = 1 To col.Count
        empIDArr(i) = col(i)(0)
        datesArr(i) = col(i)(1)
        typesArr(i) = col(i)(2)
        remarksArr(i) = col(i)(3)
    Next i
    
    Call SortFourArrays(empIDArr, datesArr, typesArr, remarksArr)
    
     '出力配列の準備
    ReDim stDate(1 To 1)
    ReDim enDate(1 To 1)
    ReDim TargetID(1 To 1)
    ReDim Types(1 To 1)
    ReDim Remarks(1 To 1)
  
        '連休開始
    curID = empIDArr(1)
    curStart = datesArr(1)
    curEnd = datesArr(1)
    curType = typesArr(1)
    curRemarks = remarksArr(1)
    
    outCnt = 1
    
    For j = 2 To UBound(datesArr)

        '同じ種別且つ同じ備考で連続している場合
        If curID = empIDArr(j - 1) And typesArr(j) = curType And remarksArr(j) = curRemarks And datesArr(j) - curEnd = 1 Then
            curEnd = datesArr(j)
        Else
            '連休終了 → 出力
            stDate(outCnt) = curStart
            enDate(outCnt) = curEnd
            TargetID(outCnt) = curID
            Types(outCnt) = curType
            Remarks(outCnt) = curRemarks

            outCnt = outCnt + 1
            ReDim Preserve stDate(1 To outCnt)
            ReDim Preserve enDate(1 To outCnt)
            ReDim Preserve TargetID(1 To outCnt)
            ReDim Preserve Types(1 To outCnt)
            ReDim Preserve Remarks(1 To outCnt)

           '新しい連休開始
            curStart = datesArr(j)
            curEnd = datesArr(j)
            curType = typesArr(j)
            curRemarks = remarksArr(j)
            curID = empIDArr(j)
        End If
    Next j
    
    outCnt = outCnt + 1
    
    ReDim Preserve stDate(1 To outCnt)
    ReDim Preserve enDate(1 To outCnt)
    ReDim Preserve TargetID(1 To outCnt)
    ReDim Preserve Types(1 To outCnt)
    ReDim Preserve Remarks(1 To outCnt)
    
    '最後の連休を出力
    stDate(outCnt) = curStart
    enDate(outCnt) = curEnd
    TargetID(outCnt) = curID
    Types(outCnt) = curType
    Remarks(outCnt) = curRemarks
   
End Sub





'バブルソート
Public Sub SortThreeArrays(ByRef eID() As Long, ByRef dates() As Date, ByRef Types() As String)
    Dim i As Long
    Dim j As Long
    Dim tempID As Long
    Dim tempDate As Date
    Dim tempType As String
    
    For i = LBound(eID) To UBound(eID) - 1
        For j = i + 1 To UBound(eID)
            
            '昇順・従業員ID
            If eID(i) > eID(j) Then
                '--- 社員IDの入れ替え---
                tempID = eID(i)
                eID(i) = eID(j)
                eID(j) = tempID
                
                '--- 日付の入れ替え ---
                tempDate = dates(i)
                dates(i) = dates(j)
                dates(j) = tempDate
                
                '--- 種別の入れ替え---
                tempType = Types(i)
                Types(i) = Types(j)
                Types(j) = tempType
                
           '従業員IDが同じ場合は日付の昇順
            ElseIf eID(i) = eID(j) And dates(i) > dates(j) Then
                '--- 社員IDの入れ替え---
                tempID = eID(i)
                eID(i) = eID(j)
                eID(j) = tempID
                
                '--- 日付の入れ替え ---
                tempDate = dates(i)
                dates(i) = dates(j)
                dates(j) = tempDate
                
                '--- 種別の入れ替え---
                tempType = Types(i)
                Types(i) = Types(j)
                Types(j) = tempType
                
            End If
        Next j
    Next i
End Sub


'バブルソート
Public Sub SortTwoArrays(ByRef datesArr() As Variant, ByRef typesArr() As Variant)
    Dim i As Long
    Dim j As Long
    Dim tempDate As Date
    Dim tempType As String

'--- バブルソートで昇順に並べ替え ---
    For i = LBound(datesArr) To UBound(datesArr) - 1
        For j = i + 1 To UBound(datesArr)

            If datesArr(i) > datesArr(j) Then
                '--- 日付の入れ替え ---
                tempDate = datesArr(i)
                datesArr(i) = datesArr(j)
                datesArr(j) = tempDate
                
                '--- 種別の入れ替え---
                tempType = typesArr(i)
                typesArr(i) = typesArr(j)
                typesArr(j) = tempType
            End If
        Next j
    Next i
End Sub


'バブルソート(CalendarFormで使用)
Public Sub SortArrayAsc(ByRef arrKeys() As Variant)
    Dim i As Long
    Dim j As Long
    Dim temp As Variant
    
    '配列が空ならExit
    If UBound(arrKeys) < LBound(arrKeys) Then Exit Sub

    '--- バブルソートで昇順に並べ替え ---
    For i = LBound(arrKeys) To UBound(arrKeys) - 1
        For j = i + 1 To UBound(arrKeys)

            If arrKeys(i) > arrKeys(j) Then
                '--- 「年」か「月」を並び替え ---
                temp = arrKeys(i)
                arrKeys(i) = arrKeys(j)
                arrKeys(j) = temp
                
            End If
        Next j
    Next i
End Sub


'バブルソート
Public Sub SortThree(ByRef datesArr() As Variant, ByRef typesArr() As Variant, ByRef remarksArr() As Variant)
    Dim i As Long
    Dim j As Long
    Dim tempDate As Date
    Dim tempType As String
    Dim tempRemarks As String

'--- バブルソートで昇順に並べ替え ---
    For i = LBound(datesArr) To UBound(datesArr) - 1
        For j = i + 1 To UBound(datesArr)

            If datesArr(i) > datesArr(j) Then
                '--- 日付の入れ替え ---
                tempDate = datesArr(i)
                datesArr(i) = datesArr(j)
                datesArr(j) = tempDate
                
                '--- 種別の入れ替え---
                tempType = typesArr(i)
                typesArr(i) = typesArr(j)
                typesArr(j) = tempType
                
                '--- 備考の入れ替え---
                tempRemarks = remarksArr(i)
                remarksArr(i) = remarksArr(j)
                remarksArr(j) = tempRemarks
            End If
        Next j
    Next i
End Sub


'バブルソート
Public Sub SortFourArrays(ByRef eID() As Long, ByRef dates() As Date, ByRef Types() As String, ByRef RM() As String)
    Dim i As Long
    Dim j As Long
    Dim tempID As Long
    Dim tempDate As Date
    Dim tempType As String
    Dim tempRemark As String
    
    For i = LBound(eID) To UBound(eID) - 1
        For j = i + 1 To UBound(eID)
            
            '昇順・従業員ID
            If eID(i) > eID(j) Then
                '--- 社員IDの入れ替え---
                tempID = eID(i)
                eID(i) = eID(j)
                eID(j) = tempID
                
                '--- 日付の入れ替え ---
                tempDate = dates(i)
                dates(i) = dates(j)
                dates(j) = tempDate
                
                '--- 種別の入れ替え---
                tempType = Types(i)
                Types(i) = Types(j)
                Types(j) = tempType
                
                '--- 備考の入れ替え---
                tempRemark = RM(i)
                RM(i) = RM(j)
                RM(j) = tempRemark
                
           '従業員IDが同じ場合は日付の昇順
            ElseIf eID(i) = eID(j) And dates(i) > dates(j) Then
                '--- 社員IDの入れ替え---
                tempID = eID(i)
                eID(i) = eID(j)
                eID(j) = tempID
                
                '--- 日付の入れ替え ---
                tempDate = dates(i)
                dates(i) = dates(j)
                dates(j) = tempDate
                
                '--- 種別の入れ替え---
                tempType = Types(i)
                Types(i) = Types(j)
                Types(j) = tempType
                
                '--- 備考の入れ替え---
                tempRemark = RM(i)
                RM(i) = RM(j)
                RM(j) = tempRemark
                
            End If
        Next j
    Next i
End Sub


