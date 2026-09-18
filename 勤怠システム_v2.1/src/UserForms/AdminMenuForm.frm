VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminMenuForm 
   ClientHeight    =   6840
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6780
   OleObjectBlob   =   "AdminMenuForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "AdminMenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Initialize
Private Sub UserForm_Initialize()
    '色
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(236, 240, 241)
    btnApproval.BackColor = RGB(52, 152, 219)
    btnApproval.ForeColor = RGB(255, 255, 255)
    btnListOfAllApplications.BackColor = RGB(52, 152, 219)
    btnListOfAllApplications.ForeColor = RGB(255, 255, 255)
    btnLeaveTypeMaster.BackColor = RGB(52, 152, 219)
    btnLeaveTypeMaster.ForeColor = RGB(255, 255, 255)
    btnEmployee.BackColor = RGB(52, 152, 219)
    btnEmployee.ForeColor = RGB(255, 255, 255)
    btnRecalc.BackColor = RGB(52, 152, 219)
    btnRecalc.ForeColor = RGB(255, 255, 255)
    btnData.BackColor = RGB(52, 152, 219)
    btnData.ForeColor = RGB(255, 255, 255)
    btnBack.BackColor = RGB(52, 152, 219)
    btnBack.ForeColor = RGB(255, 255, 255)

End Sub

'申請承認ボタン（ApproveForm へ遷移）
Private Sub btnApproval_Click()
    Me.Hide
    ApproveForm.Show
End Sub

'全申請一覧ボタン（HistoryAllForm へ遷移）
Private Sub btnListOfAllApplications_Click()
    Me.Hide
    HistoryAllForm.Show
End Sub


'休暇種別マスタボタン(Leave Typeシートをactivate)
Private Sub btnLeaveTypeMaster_Click()
    Dim res  As result
    res = Act_LT_sheet

    Select Case res
        Case InvalidInput
            MsgBox "Leave Typeシートが存在しません。", vbCritical

        Case Other
            MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
        End Select
End Sub


'社員管理ボタン（EmployeeForm へ遷移）
Private Sub btnEmployee_Click()
    Me.Hide
    EmployeeForm.Show
End Sub


'勤怠再計算ボタン（modAdmin の RecalcAllAttendance を実行）
Private Sub btnRecalc_Click()

    If MsgBox("今月分の全社員の勤怠を再計算します。よろしいですか？", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    Dim res  As result
    res = isAdmin(LoggedInID)
    
    Select Case res
        Case Success
            Dim r2  As result
            r2 = RecalcAllAttendance()
            Select Case r2
                Case Success
                    MsgBox "今月分の全社員の勤怠再計算が完了しました。", vbInformation
                    
                Case InvalidInput
                    MsgBox "今月の勤怠シートが存在しないか、再計算対象の記録がありません。", vbExclamation
                    
                Case Other
                    MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
                End Select
                        
        Case Other
            MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
    End Select
    
End Sub

'データ管理ボタン（今月のAttendance シートを開く）
Private Sub btnData_Click()

    Call Sort_Attendance
    
    Dim res  As result
    res = Act_Att_sheet
    Select Case res
        Case InvalidInput
            MsgBox "今月の勤怠シートが存在しません。", vbExclamation
        Case Other
            MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
        End Select
    
End Sub


'戻るボタン（MainMenuForm へ戻る）
Private Sub btnBack_Click()
    Me.Hide
    MainMenuForm.Show
End Sub



'フォームを閉じる（開発モードでなければ完全に閉じる）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    HandleFormClose Me, Cancel, CloseMode
End Sub

