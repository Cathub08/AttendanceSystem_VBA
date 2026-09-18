VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainMenuForm 
   ClientHeight    =   7032
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6780
   OleObjectBlob   =   "MainMenuForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MainMenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginUserID As Long

'Initialize（ログイン中ユーザー名の表示 & 権限チェック）
Private Sub UserForm_Initialize()

    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(236, 240, 241)
    
    btnLogout.BackColor = RGB(52, 152, 219)
    btnLogout.ForeColor = RGB(255, 255, 255)
    btnClock.BackColor = RGB(52, 152, 219)
    btnClock.ForeColor = RGB(255, 255, 255)
    btnLeave.BackColor = RGB(52, 152, 219)
    btnLeave.ForeColor = RGB(255, 255, 255)
    btnList.BackColor = RGB(52, 152, 219)
    btnList.ForeColor = RGB(255, 255, 255)
    btnReport.BackColor = RGB(52, 152, 219)
    btnReport.ForeColor = RGB(255, 255, 255)
    btnChangePassword.BackColor = RGB(52, 152, 219)
    btnChangePassword.ForeColor = RGB(255, 255, 255)
    btnAdmin.BackColor = RGB(52, 152, 219)
    btnAdmin.ForeColor = RGB(255, 255, 255)
    
    
End Sub

'LoginUserID を使う処理
Private Sub UserForm_Activate()

    If LoginUserID = 0 Then Exit Sub   'ログイン前は何もしない

    btnAdmin.Visible = IsAdminByID(LoginUserID)
    
    Call LoadUserInfo
    
End Sub


Private Sub LoadUserInfo()
    lblUser.Caption = "ログイン中：" & GetLoggedInName() & "（ID：" & LoginUserID & "）"
End Sub


'出退勤打刻画面へ遷移
Private Sub btnClock_Click()
    Me.Hide
    ClockForm.Show
End Sub

'休暇申請画面へ遷移
Private Sub btnLeave_Click()
    Me.Hide
    LeaveForm.Show
End Sub

'勤怠一覧表示（ShowAttendanceList）
Private Sub btnList_Click()

    Dim res  As result
    res = ShowAttendanceList
        Select Case res
            Case Other
                MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
        End Select
        
End Sub

'月次レポート出力
Private Sub btnReport_Click()

    Call Sort_Attendance
    
    Dim res  As result
    res = GenerateMonthlyReport
    Select Case res
        Case Success
            MsgBox "月次レポートを作成しました。", vbInformation

            'csv出力を選択
            If MsgBox("月次レポートをCSV出力しますか？", vbYesNo + vbQuestion) = vbYes Then
                Call ExportMonthlyReportCSV
            End If
        Case Other
            MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
    End Select
        
End Sub

'パスワード変更画面へ遷移
Private Sub btnChangePassword_Click()
    Me.Hide
    ChangePasswordForm.Show
End Sub

'管理者メニューへ遷移
Private Sub btnAdmin_Click()
    Me.Hide
    AdminMenuForm.Show
    
End Sub

'ログアウト
Private Sub btnLogout_Click()
    'Excel感を戻す
    RestoreExcelLook

    'Mainシートを表示
    If SheetExists(SHEET_MAIN) Then
        ThisWorkbook.Sheets(SHEET_MAIN).Activate
    End If

    Logout
    Me.Hide
    LoginForm.Show
End Sub


'フォームを閉じる（開発モードでなければ完全に閉じる）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    HandleFormClose Me, Cancel, CloseMode
End Sub
