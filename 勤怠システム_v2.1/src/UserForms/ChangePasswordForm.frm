VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangePasswordForm 
   ClientHeight    =   6240
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5988
   OleObjectBlob   =   "ChangePasswordForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ChangePasswordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Initialize（フォーム起動時の初期化）
Private Sub UserForm_Initialize()
    '色
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)
    btnRegistration.BackColor = RGB(52, 152, 219)
    btnRegistration.ForeColor = RGB(255, 255, 255)
    btnBack.BackColor = RGB(52, 152, 219)
    btnBack.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(236, 240, 241)

    '入力欄初期化
    txtOldPass.value = ""
    txtNewPass.value = ""
    txtConfirmPass.value = ""

    'メッセージ消去
    lblMessage.Caption = ""

End Sub


'登録ボタン
Private Sub btnRegistration_Click()
    Dim empID As Long
    Dim pwd As String
    Dim NewPwd As String
    
    empID = GetLoggedInID()   'ログイン中の従業員IDをempIDに入れる
    pwd = txtOldPass.value    '旧パスワードの入力値を格納
    NewPwd = txtNewPass.value '新パスワードの入力値を格納

    '入力チェック(空)
    If txtOldPass.value = "" Or txtNewPass.value = "" Or txtConfirmPass.value = "" Then
        lblMessage.Caption = "パスワードを入力してください。"
        Exit Sub
    End If
    
    '入力チェック(英数字8桁)
    If Not IsValidPwd(NewPwd) Then
        lblMessage.Caption = "パスワードは英数字混在の8桁で設定してください"
        Exit Sub
    End If

    '入力チェック(新PWと新PW(確認)の一致)
    If Trim(txtNewPass.value) <> Trim(txtConfirmPass.value) Then
        lblMessage.Caption = "新パスワードと新パスワード(確認)の入力が不一致です"
        Exit Sub
    End If
    
    'パスワード更新処理
    Dim res  As result
    res = UpdatePassword(empID, pwd, NewPwd)
    Select Case res
        Case Success
            lblMessage.Caption = "登録が完了しました"
            txtOldPass.value = ""
            txtNewPass.value = ""
            txtConfirmPass.value = ""
        Case InvalidInput
            lblMessage.Caption = "旧パスワードの入力が間違っています。"
        Case Other
            lblMessage.Caption = "システム内部のエラーにより処理を完了できませんでした｡"
        End Select
        
End Sub



'戻るボタン
Private Sub btnBack_Click()
    Unload Me
    MainMenuForm.Show
End Sub


'フォームを閉じる（開発モードでなければ完全に閉じる）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    HandleFormClose Me, Cancel, CloseMode
End Sub


