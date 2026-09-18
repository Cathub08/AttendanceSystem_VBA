VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LeaveForm 
   ClientHeight    =   8040
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7980
   OleObjectBlob   =   "LeaveForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LeaveForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub UserForm_Initialize()

    '色
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(236, 240, 241)
    btnSubmit.BackColor = RGB(52, 152, 219)
    btnSubmit.ForeColor = RGB(255, 255, 255)
    btnBack.BackColor = RGB(52, 152, 219)
    btnBack.ForeColor = RGB(255, 255, 255)
    btnUndohistory.BackColor = RGB(52, 152, 219)
    btnUndohistory.ForeColor = RGB(255, 255, 255)
        
    '申請者名をラベルへ表示
     Dim empName As String
     empName = GetLoggedInName()
    lblApplicant.Caption = "申請者：" & empName

    'テキストボックスの値を初期化
    If txtStartDate.value <> "" Then
        txtStartDate.value = ""
    End If

    If txtEndDate.value <> "" Then
        txtEndDate.value = ""
    End If

End Sub

'コンボボックスへ初期値「休暇種別」をセット
Private Sub UserForm_Activate()

    'シート存在チェック
    If Not SheetExists(SHEET_LEAVE) Then
        MsgBox "休暇種別シートが存在しないため、申請できません。", vbExclamation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets(SHEET_LEAVE)
    lr = GetLastRow(ws, LT_COL_LEAVE_TYPE)
    
    '休暇種別をセット
    cmbType.Clear
    For i = 2 To lr
        If Trim(ws.Cells(i, LT_COL_LEAVE_TYPE).value) <> "" Then
            cmbType.AddItem ws.Cells(i, LT_COL_LEAVE_TYPE).value
        End If
    Next i
    
    '手入力禁止、リストからの選択のみ許可
    cmbType.Style = fmStyleDropDownList
    
    If cmbType.ListCount = 0 Then
        MsgBox "休暇種別マスタが空のため、申請できません。", vbExclamation
        Exit Sub
    End If
    
End Sub


'①休暇開始日ボタン
Private Sub btnStart_Click()
    If txtStartDate.value <> "" Then
        txtStartDate.value = ""
    End If
    Unload CalendarForm
    CalendarForm.CalendarMode = "Future3"
    Call CalendarForm.Mode_Future3
    Set CalendarForm.TargetTextbox = Me.txtStartDate
    CalendarForm.Show
End Sub


'②休暇終了日ボタン
Private Sub btnEnd_Click()
    If txtEndDate.value <> "" Then
        txtEndDate.value = ""
    End If
    Unload CalendarForm
    CalendarForm.CalendarMode = "Future3"
    Call CalendarForm.Mode_Future3
    Set CalendarForm.TargetTextbox = Me.txtEndDate
    CalendarForm.Show
End Sub


'③登録ボタン（休暇申請登録）
Private Sub btnSubmit_Click()
    Dim empID As Long
    Dim leaveType As String
    Dim note As String
    Dim STARTDate As Date
    Dim ENDDate As Date
    empID = GetLoggedInID()

    '入力チェック
    '開始日・終了日のどちらか空欄ならエラー
    If txtStartDate.value = "" Or txtEndDate.value = "" Then
        MsgBox "開始日と終了日を入力してください。", vbExclamation
        Exit Sub
    End If

    '開始日に終了日よりも大きい数字が設定されていたらエラー
    If CDate(txtStartDate.value) > CDate(txtEndDate.value) Then
        MsgBox "終了日は開始日以降の日付を選択してください。", vbExclamation
        Exit Sub
    End If

    '開始日が今日より前ならエラー
    If CDate(txtStartDate.value) < Date Then
        MsgBox "開始日は今日以降の日付を選択してください。", vbExclamation
        Exit Sub
    End If

    '休暇種別が未選択ならエラー
    If cmbType.value = "" Then
        MsgBox "休暇種別を選択してください。", vbExclamation
        Exit Sub
    End If

    '標準モジュールで参照の為、変数へ格納
    leaveType = cmbType.value
    note = txtNote.value
    STARTDate = txtStartDate.value
    ENDDate = txtEndDate.value
    
    '休暇登録（modClock の RegisterLeave を使用）
    Dim res  As result
    res = RegisterLeave(empID, STARTDate, ENDDate, leaveType, note)
    
    Select Case res
        Case Success
            MsgBox "休暇申請を登録しました。", vbInformation
            Unload Me
            MainMenuForm.Show
            
        Case Other
            MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
        End Select

End Sub


'④申請履歴・取消ボタン
Private Sub btnUndohistory_Click()
    Unload Me
    HistoryForm.Show
End Sub


'⑤戻るボタン
Private Sub btnBack_Click()
    Unload Me
    MainMenuForm.Show
End Sub


'フォームを閉じる（開発モードでなければ完全に閉じる）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    HandleFormClose Me, Cancel, CloseMode
End Sub


