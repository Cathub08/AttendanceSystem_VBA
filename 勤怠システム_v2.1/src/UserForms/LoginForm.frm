VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   ClientHeight    =   5232
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4584
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit


'① Initialize（フォーム起動時の初期化）
Private Sub UserForm_Initialize()

    'タイトル設定
    lblTitle.Caption = "勤怠管理システム"

    'タイトルの色（濃紺）
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)

    'ログインボタンの色（青）
    btnLogin.BackColor = RGB(52, 152, 219)
    btnLogin.ForeColor = RGB(255, 255, 255)

    'フォーム背景（薄いグレー）
    Me.BackColor = RGB(236, 240, 241)

    '入力欄初期化
    txtID.value = ""
    txtPass.value = ""

    'メッセージ消去
    lblMessage.Caption = ""

    'Attendanceシートの存在チェックとなければ追加【↓Test時はコメントアウトしないと、未来分を含んだATTシート無しの状態がつくれない】
    Call SheetExistsAndGetFutureSheet
    
End Sub



'② ログインボタン（btnLogin）クリック時のコード
Private Sub btnLogin_Click()
    Dim empID As Long
    Dim pwd As String
    Dim res  As result
    
    '入力チェック
    If txtID.value = "" Or txtPass.value = "" Then
        lblMessage.Caption = "社員IDとパスワードを入力してください。"
        Exit Sub
    End If

    If Not IsNumeric(txtID.value) Then
        lblMessage.Caption = "社員IDは数字で入力してください。"
        Exit Sub
    End If
    
    empID = CLng(txtID.value)
    pwd = txtPass.value
        
    'ID存在確認
    res = EmployeeExists(empID)
    Select Case res
        Case InvalidInput
            lblMessage.Caption = "社員IDが存在しません。"
            Exit Sub
        Case Other
            lblMessage.Caption = "システム内部のエラーにより処理を完了できませんでした｡"
            Exit Sub
    End Select
       
       
    'ログイン処理
    res = Login(empID, pwd)

    Select Case res
        Case Success
            LoggedInID = CInt(txtID.text) 'modLoginのPublic変数[LoggedInID]へ格納
            'MsgBox "LoggedInIDは" & LoggedInID
            
            MainMenuForm.LoginUserID = LoggedInID
            'MsgBox "MainMenuForm.LoginUserIDは" & MainMenuForm.LoginUserID
                     
            Unload Me
            
            'Excel感を消す
            RemoveExcelLook
            'Mainシートを表示
            If SheetExists(SHEET_MAIN) Then
                ThisWorkbook.Sheets(SHEET_MAIN).Activate
            End If
            
            MainMenuForm.Show
            
        Case InvalidInput
            lblMessage.Caption = "ID またはパスワードが違います"
            
        Case Other
            lblMessage.Caption = "システム内部のエラーにより処理を完了できませんでした｡"
        End Select
        
End Sub


'フォームを閉じる（開発モードでなければ完全に閉じる）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    HandleFormClose Me, Cancel, CloseMode
End Sub


