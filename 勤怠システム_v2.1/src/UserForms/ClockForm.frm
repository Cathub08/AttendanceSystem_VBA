VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClockForm 
   ClientHeight    =   6840
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7584
   OleObjectBlob   =   "ClockForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ClockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'① Initialize
Private Sub UserForm_Initialize()

    'タイトル色
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)

    '背景色
    Me.BackColor = RGB(236, 240, 241)
    
    'ボタンの色（青）
    btnBack.BackColor = RGB(52, 152, 219)
    btnBack.ForeColor = RGB(255, 255, 255)
    btnIn.BackColor = RGB(52, 152, 219)
    btnIn.ForeColor = RGB(255, 255, 255)
    btnOut.BackColor = RGB(52, 152, 219)
    btnOut.ForeColor = RGB(255, 255, 255)
    btnBreakRegistration.BackColor = RGB(52, 152, 219)
    btnBreakRegistration.ForeColor = RGB(255, 255, 255)

    '今日の日付
    lblDate.Caption = "日付：" & Format(Date, "yyyy/mm/dd")
    
    '入力欄初期化
    txtBreak_time.value = ""
    
    Call LoadTodayAttendance
End Sub
    
    
'今日の出退勤情報をUIへ反映
Private Sub LoadTodayAttendance()
    '今日の打刻状況を取得
    Dim cin As Variant, cout As Variant, Break As Variant
    cin = GetClockIn(GetLoggedInID(), Date)
    cout = GetClockOut(GetLoggedInID(), Date)
    Break = GetBreakTime(GetLoggedInID(), Date)
    
    '出勤時刻表示
    If IsEmpty(cin) Or cin = "" Then
        lblClockIn.Caption = "出勤：--:--"
    Else
        lblClockIn.Caption = "出勤：" & Format(cin, "hh:mm")
        btnIn.Enabled = False '出勤済みならボタン無効
    End If

    '退勤時刻表示
    If IsEmpty(cout) Or cout = "" Then
        lblClockOut.Caption = "退勤：--:--"
    Else
        lblClockOut.Caption = "退勤：" & Format(cout, "hh:mm")
        btnOut.Enabled = False '退勤済みならボタン無効
    End If

    '休憩時間表示
     If IsEmpty(Break) Or Break = "" Then
        lblShowBreakTime.Caption = "休憩：--:--"
    Else
        lblShowBreakTime.Caption = "休憩：" & Format(Break, "hh:mm")
        '休憩時間入力済の場合の処理を追加する際はここに記載
    End If

End Sub

'② 出勤打刻ボタン
Private Sub btnIn_Click()
    Dim res  As result
    res = ClockIn(GetLoggedInID(), Date)
    
    Select Case res
        Case Success
            MsgBox "出勤打刻に成功しました。", vbInformation
            
        Case InvalidInput
            MsgBox "本日はすでに出勤済みです。", vbExclamation
            
        Case Other
            MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
        End Select

    Call LoadTodayAttendance

End Sub


'③ 退勤打刻ボタン
Private Sub btnOut_Click()

    Dim res  As result
    res = ClockOut(GetLoggedInID(), Date)

    Select Case res
        Case Success
            MsgBox "退勤打刻に成功しました。", vbInformation
            
        Case InvalidInput
            MsgBox "本日はすでに退勤済みです。", vbExclamation
        
        Case NoRecordsForToday
            MsgBox "本日の出勤記録がありません。", vbExclamation
        
        Case Other
            MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
    End Select
           
    Call LoadTodayAttendance

End Sub


' 休憩登録ボタン
Private Sub btnBreakRegistration_Click()
    Dim BREAK_TIME As Date

    '入力チェック
    On Error Resume Next
    BREAK_TIME = TimeValue(txtBreak_time.value)
    If Err.Number <> 0 Then
        MsgBox "休憩時間は 0:00 の形式で入力してください。", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    'シートへ反映
    BREAK_TIME = CDate(txtBreak_time.value)
    
    Dim res  As result
    res = ReflectBreakTimes(GetLoggedInID(), Date, BREAK_TIME)
    
    Select Case res
        Case Success
            MsgBox "休憩時間の反映に成功しました。", vbInformation
        Case NoRecordsForToday
            MsgBox "本日の出勤記録がありません。", vbExclamation
        Case Other
            MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
        End Select
    
    Call LoadTodayAttendance
End Sub



' 戻るボタン
Private Sub btnBack_Click()
    Unload Me
    MainMenuForm.Show
End Sub



'フォームを閉じる（開発モードでなければ完全に閉じる）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    HandleFormClose Me, Cancel, CloseMode
End Sub


