VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   ClientHeight    =   7836
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7188
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim futureDate As Collection
Public CalendarMode As String
Public TargetTextbox As Variant
Public TestDate As Date
Dim Dict As Object
Dim DictMonth As Object
    
Private Sub UserForm_Initialize()
    '配色設定
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(236, 240, 241)
    btnToday.BackColor = RGB(52, 152, 219)
    btnToday.ForeColor = RGB(255, 255, 255)
    btnClose.BackColor = RGB(52, 152, 219)
    btnClose.ForeColor = RGB(255, 255, 255)
    
    '初期化
    cboYear.Clear
    cboMonth.Clear
    

End Sub


'-----呼出Fromによってモード切替-----
Private Sub UserForm_Activate()
    Select Case CalendarMode
        Case "AllApplications" 'HistoryAllFormの場合
            Mode_AllApplications
    End Select
    
'    'TEST↓------------------
'    Debug.Print "今開いているModeは" & CalendarMode
'    Debug.Print "TargetTextboxは" & TargetTextbox.Parent.name & "." & TargetTextbox.name
'    'TEST↑------------------
    
End Sub


'呼出元が[LeaveForm]の場合の「年月辞書化」
Public Sub Mode_Future3()
    Set Dict = CreateObject("Scripting.Dictionary")
    Set DictMonth = CreateObject("Scripting.Dictionary")
    
    Dim d As Date, i As Long
    Dim y As Variant, m As Variant
    
    'TestDate = DateSerial(2026, 11, 1)
    'TestDate = DateSerial(2026, 8, 1)
    TestDate = Date
    d = TestDate
    
    For i = 0 To 3
        
        y = CStr(year(DateAdd("m", i, d)))
        m = Format(month(DateAdd("m", i, d)), "00")
        
        '（年→月）辞書
        If Not Dict.Exists(y) Then
            Dict.Add y, Array(m)
        Else
            '既にある配列に、新しい月 m を追加(これをしないとKeyは重複不可のため、年重複時の月が欠落する)
            Dict(y) = AddToArray(Dict(y), m)
        End If

        '（月→年）辞書
        If Not DictMonth.Exists(m) Then
            DictMonth.Add m, y
        End If
    Next i
    
    
'    'TESTで使用。辞書化確認↓-------------------------------------------------------------------
'    'Dict確認
'    Dim k As Variant
'    Dim arr As Variant
'    Dim elem As Variant
'
'    For Each k In Dict.Keys
'        arr = Dict(k)
'        For Each elem In arr
'            'Debug.Print "辞書Dict格納値：" & k & "：" & elem
'            Debug.Print "Dict.Keys →" & k
'            Debug.Print "Dict(" & k & ")→" & elem
'        Next elem
'    Next k
'
'    'DictMonth確認
'    Dim km As Variant
'
'    For Each km In DictMonth.Keys
'        'Debug.Print "辞書DictMonth格納値：" & km & "：" & DictMonth(km)
'        Debug.Print "DictMonth.Keys →" & km
'        Debug.Print "DictMonth(" & km & ")→" & DictMonth(km)
'    Next km
'
'    'TESTで使用。ここまで↑---------------------------------------------------------------------
    
End Sub


'呼出元が[HistoryAllForm]の場合の「年月コンボボックス」の設定
Private Sub Mode_AllApplications()
    Dim y As Long, m As Long

    cboYear.Clear
    cboMonth.Clear
    
    ' 年の設定（2000年-今日を基準に5年後まで）
    For y = 2000 To year(Date) + 5
        cboYear.AddItem y
    Next y

    '月を設定
    For m = 1 To 12
        cboMonth.AddItem Format(m, "00")
    Next m
    ' 手入力禁止、リストからの選択のみ許可
    cboYear.Style = fmStyleDropDownList
    cboMonth.Style = fmStyleDropDownList

End Sub



Private Sub cboYear_Enter()

    'AllApplications モードならカレンダーセットのみ
    If CalendarMode = "AllApplications" Then
        If cboYear.value <> "" And cboMonth.value <> "" Then
            Call CalendarSet
        End If
        Exit Sub
    End If
    
    'Future3モードの処理
    cboYear.Clear

    Dim arrYears() As Variant
    Dim arrMonths As Variant
    Dim y As Variant
    Dim m As Variant

    '月が未選択 → 全年を表示
    If cboMonth.value = "" Then
        arrYears = Dict.Keys
        Call SortArrayAsc(arrYears)

        For Each y In arrYears
            cboYear.AddItem y
        Next y

    Else
        '月が選択済み → 月→年辞書から紐づく年だけ
        cboYear.AddItem DictMonth(cboMonth.value)
    End If
    
    cboYear.Style = fmStyleDropDownList
    
End Sub


Private Sub cboMonth_Enter()

    'AllApplications モードならカレンダーセットのみ
    If CalendarMode = "AllApplications" Then
        If cboYear.value <> "" And cboMonth.value <> "" Then
            Call CalendarSet
        End If
        Exit Sub
    End If
    
    'Future3モードの処理
    cboMonth.Clear

    'Dim arrMonths As Variant
    Dim y As Variant
    Dim m As Variant

    '年が未選択 → 全月を表示
    If cboYear.value = "" Then
        '全月を集める
        Dim tmp As Variant
        Dim allMonths() As Variant
        Dim i As Long

        tmp = Dict.Keys

        '全月を配列にまとめる
        Dim list As Collection
        Set list = New Collection

        For Each y In tmp
            For Each m In Dict(y)
                list.Add m
            Next m
        Next y

        'Collection → 配列
        ReDim allMonths(0 To list.Count - 1)
        For i = 1 To list.Count
            allMonths(i - 1) = list(i)
        Next i

        Call SortArrayAsc(allMonths)

        For Each m In allMonths
            cboMonth.AddItem m
        Next m

    Else
        '年が選択済み → 年→月辞書から紐づく月だけ
        Dim arr() As Variant
        arr = Dict(cboYear.value)
        Call SortArrayAsc(arr)

        For Each m In arr
            cboMonth.AddItem m
        Next m
    End If

    cboMonth.Style = fmStyleDropDownList
    
End Sub


Private Sub cboYear_Change()

    If CalendarMode = "AllApplications" Then
        If cboYear.value <> "" And cboMonth.value <> "" Then
            CalendarSet
        End If
        Exit Sub
    End If

    If cboYear.value <> "" And cboMonth.value <> "" Then
        CalendarSet
    End If
    
    If cboYear.value = "" Or cboMonth.value = "" Then
        Call ClearCalendar
        Exit Sub
    End If

End Sub


Private Sub cboMonth_Change()

    If CalendarMode = "AllApplications" Then
        If cboYear.value <> "" And cboMonth.value <> "" Then
            CalendarSet
        End If
        Exit Sub
    End If

    If cboYear.value <> "" And cboMonth.value <> "" Then
        CalendarSet
    End If

    If cboYear.value = "" Or cboMonth.value = "" Then
        Call ClearCalendar
        Exit Sub
    End If

End Sub


'年月の選択に応じてカレンダーをセット
Private Sub CalendarSet()
    Dim first_day_of_the_month As Date
    Dim last_day_of_the_month As Long
    Dim The_day_of_the_week_on_the_first_day_of_the_month As Long
    Dim i As Long
    Dim InputYear As String
    Dim InputMonth As String
    Dim shoki As Long
    
    InputYear = cboYear.value
    InputMonth = cboMonth.value
    
    If InputYear = "" Or InputMonth = "" Then
    Exit Sub
    End If
    
    'カレンダー初期化
    For shoki = 1 To 42
        Me.Controls("cmdDay" & shoki).Caption = ""
    Next shoki
    
    '月初日
    first_day_of_the_month = DateSerial(InputYear, InputMonth, 1)
    '月末
    last_day_of_the_month = Format(WorksheetFunction.EoMonth(first_day_of_the_month, 0), "d")
    
    '月初日の曜日'返却値は「1」→日曜、「7」→土曜。
    The_day_of_the_week_on_the_first_day_of_the_month = Weekday(first_day_of_the_month)
    
    ' 日付をセット
    For i = 1 To last_day_of_the_month
        With Me.Controls("cmdDay" & (The_day_of_the_week_on_the_first_day_of_the_month + i) - 1)
            .Caption = i
            .Enabled = True
        End With
    Next i
    
End Sub

'日付ボタン初期化
Public Sub ClearCalendar()
    Dim i As Long
   
    For i = 1 To 42
        With Me.Controls("cmdDay" & i)
            .Caption = ""
            .Enabled = False
        End With
    Next i

End Sub


'1-42の日付ボタンを押下
Private Sub cmdDay1_Click(): DayButton_Click 1: End Sub
Private Sub cmdDay2_Click(): DayButton_Click 2: End Sub
Private Sub cmdDay3_Click(): DayButton_Click 3: End Sub
Private Sub cmdDay4_Click(): DayButton_Click 4: End Sub
Private Sub cmdDay5_Click(): DayButton_Click 5: End Sub
Private Sub cmdDay6_Click(): DayButton_Click 6: End Sub
Private Sub cmdDay7_Click(): DayButton_Click 7: End Sub
Private Sub cmdDay8_Click(): DayButton_Click 8: End Sub
Private Sub cmdDay9_Click(): DayButton_Click 9: End Sub
Private Sub cmdDay10_Click(): DayButton_Click 10: End Sub
Private Sub cmdDay11_Click(): DayButton_Click 11: End Sub
Private Sub cmdDay12_Click(): DayButton_Click 12: End Sub
Private Sub cmdDay13_Click(): DayButton_Click 13: End Sub
Private Sub cmdDay14_Click(): DayButton_Click 14: End Sub
Private Sub cmdDay15_Click(): DayButton_Click 15: End Sub
Private Sub cmdDay16_Click(): DayButton_Click 16: End Sub
Private Sub cmdDay17_Click(): DayButton_Click 17: End Sub
Private Sub cmdDay18_Click(): DayButton_Click 18: End Sub
Private Sub cmdDay19_Click(): DayButton_Click 19: End Sub
Private Sub cmdDay20_Click(): DayButton_Click 20: End Sub
Private Sub cmdDay21_Click(): DayButton_Click 21: End Sub
Private Sub cmdDay22_Click(): DayButton_Click 22: End Sub
Private Sub cmdDay23_Click(): DayButton_Click 23: End Sub
Private Sub cmdDay24_Click(): DayButton_Click 24: End Sub
Private Sub cmdDay25_Click(): DayButton_Click 25: End Sub
Private Sub cmdDay26_Click(): DayButton_Click 26: End Sub
Private Sub cmdDay27_Click(): DayButton_Click 27: End Sub
Private Sub cmdDay28_Click(): DayButton_Click 28: End Sub
Private Sub cmdDay29_Click(): DayButton_Click 29: End Sub
Private Sub cmdDay30_Click(): DayButton_Click 30: End Sub
Private Sub cmdDay31_Click(): DayButton_Click 31: End Sub
Private Sub cmdDay32_Click(): DayButton_Click 32: End Sub
Private Sub cmdDay33_Click(): DayButton_Click 33: End Sub
Private Sub cmdDay34_Click(): DayButton_Click 34: End Sub
Private Sub cmdDay35_Click(): DayButton_Click 35: End Sub
Private Sub cmdDay36_Click(): DayButton_Click 36: End Sub
Private Sub cmdDay37_Click(): DayButton_Click 37: End Sub
Private Sub cmdDay38_Click(): DayButton_Click 38: End Sub
Private Sub cmdDay39_Click(): DayButton_Click 39: End Sub
Private Sub cmdDay40_Click(): DayButton_Click 40: End Sub
Private Sub cmdDay41_Click(): DayButton_Click 41: End Sub
Private Sub cmdDay42_Click(): DayButton_Click 42: End Sub


'日付ボタンを押下
Private Sub DayButton_Click(Index As Integer)
    Dim tb As MSForms.TextBox
    
    On Error Resume Next ' エラーを無視して処理を続行
    If TargetTextbox Is Nothing Then
    On Error GoTo 0      ' エラー処理を元に戻す
        MsgBox "転記先のテキストボックスが設定されていません。", vbExclamation
        Exit Sub
    End If

    Set tb = TargetTextbox
    tb.value = DateSerial(cboYear.value, cboMonth.value, Me.Controls("cmdDay" & Index).Caption)
    Unload Me
End Sub



'今日へボタン
Private Sub btnToday_Click()

    Dim ThisM As Date
    Dim d As Date
    
    Select Case CalendarMode
        Case "Future3" 'LeaveFormの場合
            
            d = TestDate
            ThisM = month(d)
            
            Call SetCmb
            
            '今日の年月を設定
            cboYear.value = year(d)
            cboMonth.value = Format(ThisM, "00")
                 
            Call CalendarSet
            
        Case "AllApplications" 'HistoryAllFormの場合
            d = Date
            ThisM = month(d)
            
             '今日の年月を設定
            cboYear.value = year(d)
            cboMonth.value = Format(ThisM, "00")
                 
            Call CalendarSet
    End Select
End Sub


'閉じるボタン
Private Sub btnClose_Click()
    Unload Me
End Sub


'選択できるようにプルダウンに全選択肢を入れなおす。↓
Public Sub SetCmb()
    Dim arrYears() As Variant
    Dim y As Variant
    
    '●cboYear_Add '年プルダウンAdd
    arrYears = Dict.Keys
    Call SortArrayAsc(arrYears)

    cboYear.value = ""
    
    For Each y In arrYears
        cboYear.AddItem y
    Next y

    Dim m As Variant
    Dim tmp As Variant
    Dim allMonths() As Variant
    Dim i As Long
        
    '●cboMonth_Add '月プルダウンAdd
    tmp = Dict.Keys

    '全月を配列にまとめる
    Dim list As Collection
    Set list = New Collection

    For Each y In tmp
        For Each m In Dict(y)
            list.Add m
        Next m
    Next y

    'Collection → 配列
    ReDim allMonths(0 To list.Count - 1)
    For i = 1 To list.Count
        allMonths(i - 1) = list(i)
    Next i

    Call SortArrayAsc(allMonths)

    cboMonth.value = ""
    
    For Each m In allMonths
        cboMonth.AddItem m
    Next m
        
End Sub



