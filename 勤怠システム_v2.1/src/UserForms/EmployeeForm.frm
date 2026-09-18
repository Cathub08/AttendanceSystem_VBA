VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmployeeForm 
   ClientHeight    =   9732.001
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9864.001
   OleObjectBlob   =   "EmployeeForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "EmployeeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurrentMode As String


Private Sub UserForm_Initialize()
    '色
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(236, 240, 241)
    btnBack.BackColor = RGB(52, 152, 219)
    btnBack.ForeColor = RGB(255, 255, 255)
    btnAdd.BackColor = RGB(52, 152, 219)
    btnAdd.ForeColor = RGB(255, 255, 255)
    btnUpdate.BackColor = RGB(52, 152, 219)
    btnUpdate.ForeColor = RGB(255, 255, 255)
    btnSearch.BackColor = RGB(52, 152, 219)
    btnSearch.ForeColor = RGB(255, 255, 255)
    btnDelete.BackColor = RGB(52, 152, 219)
    btnDelete.ForeColor = RGB(255, 255, 255)
    lblOperationGuide.BackColor = RGB(255, 255, 200)
    
    lstEmployees.Clear
    lstEmployees.ColumnCount = 2
    lstEmployees.ColumnWidths = "50 pt;120 pt"
    lstEmployees.Height = 315
    lblOperationGuide.Caption = ""
    lblOperationGuide.Caption = "はじめに更新・追加・削除ボタンから作業を選択してください"
        
    Call ResetToInitialMode
    
End Sub
    
Private Sub UserForm_Activate()
    Call LoadDepartments
    Call LoadEmployeeList
End Sub
    
'コンボボックスへ部門登録
Private Sub LoadDepartments()

    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long

    'シート存在チェック
    If Not SheetExists(SHEET_DEP) Then
        MsgBox "部門マスタが存在しません。" & vbCrLf & "管理者に連絡してください。", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_DEP) ' 部門一覧シート

    lr = GetLastRow(ws, DEP_COL_DEPARTMENT)

    cmbDep.Clear
    For i = 2 To lr
        If Trim(ws.Cells(i, DEP_COL_DEPARTMENT).value) <> "" Then
            cmbDep.AddItem ws.Cells(i, DEP_COL_DEPARTMENT).value
        End If
    Next i
    '手入力禁止、リストからの選択のみ許可
     cmbDep.Style = fmStyleDropDownList
End Sub



' 従業員一覧を読み込む専用プロシージャ
Private Sub LoadEmployeeList()

    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long

    'シート存在チェック
    If Not SheetExists(SHEET_EMP) Then
        MsgBox "社員マスタが存在しません。" & vbCrLf & "管理者に連絡してください。", vbExclamation
        Exit Sub
    End If
    
    lstEmployees.Clear
    
    Set ws = ThisWorkbook.Sheets(SHEET_EMP)

    lr = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)

    For i = 2 To lr
        lstEmployees.AddItem
        lstEmployees.list(lstEmployees.ListCount - 1, 0) = ws.Cells(i, EMP_COL_EMPLOYEE_ID).value
        lstEmployees.list(lstEmployees.ListCount - 1, 1) = ws.Cells(i, EMP_COL_NAME).value
    Next i
    
End Sub


'ListBox 選択時にフォームへ反映
Private Sub lstEmployees_Click()

    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long
    Dim selectedID As Long
    Dim arr
    Dim FullName As String
    Dim Parts As Variant
    
    'ListBox の選択値は「ID : 氏名」の形式なので分割
    arr = Split(lstEmployees.value, " : ")
    selectedID = CLng(lstEmployees.list(lstEmployees.ListIndex, 0))

    Set ws = ThisWorkbook.Sheets(SHEET_EMP)
    lr = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)

    'Employees シートから該当IDを検索
    For i = 2 To lr
        If ws.Cells(i, EMP_COL_EMPLOYEE_ID).value = selectedID Then
        
            FullName = ws.Cells(i, EMP_COL_NAME).value
            Parts = Split(FullName, " ")
            ' --- フォームに反映 ---
            txtLastName.value = Parts(0)
            txtFirstName.value = Parts(1)
            TextEmpId.value = ws.Cells(i, EMP_COL_EMPLOYEE_ID).value
            cmbDep.value = ws.Cells(i, EMP_COL_DEPARTMENT).value
            lblPass.Caption = "Password：" & ws.Cells(i, EMP_COL_PASSWORD).value
            chkAdmin.value = ws.Cells(i, EMP_COL_IS_ADMIN).value

            ' --- 行番号を保持（更新・削除で使用） ---
            Me.Tag = i

            Exit For
        End If
    Next i
    
'    'TEST↓------------------
'    Debug.Print "(Me.Tag)行番号:"; Me.Tag
'    'TEST↑------------------

End Sub

'初期状態にリセット
Private Sub ResetToInitialMode()
    CurrentMode = "" ' モードを未選択に戻す
    Me.Tag = ""   ' 検索で保持した行番号もクリア

    ' --- 入力欄のクリア ---
    txtLastName.value = ""
    txtFirstName.value = ""
    TextEmpId.value = ""
    cmbDep.value = ""
    lblPass.Caption = ""
    chkAdmin.value = False

    ' --- ListBox のクリア ---
    lstEmployees.Clear

    ' --- 全コントロールを無効化 ---
    txtLastName.Enabled = False
    txtFirstName.Enabled = False
    TextEmpId.Enabled = False
    cmbDep.Enabled = False
    chkAdmin.Enabled = False
    lstEmployees.Enabled = False
    btnSearch.Enabled = False

    ' --- モード切替ボタンだけ有効化 ---
    btnAdd.Enabled = True
    btnUpdate.Enabled = True
    btnDelete.Enabled = True
    
    '---　menuメッセージを表示　---
    lblOperationGuide.Caption = ""
    lblOperationGuide.Caption = "はじめに更新・追加・削除ボタンから作業を選択してください"
    
'    'TEST↓------------------
'    Debug.Print "今のModeは" & CurrentMode
'    Debug.Print "(Me.Tag)行番号:"; Me.Tag
'    'TEST↑------------------

End Sub

'【SetMode_Add（追加モード）】
Private Sub SetMode_Add()

    CurrentMode = "Add"

    lblOperationGuide.Caption = ""
    lblOperationGuide.Caption = "追加する氏名・部門名・管理者区分を入れて追加ボタンを押してください"
        
    ' --- 入力欄の有効化 ---
    txtLastName.Enabled = True
    txtFirstName.Enabled = True
    cmbDep.Enabled = True
    chkAdmin.Enabled = True

    ' ID は自動採番なので入力不可
    TextEmpId.Enabled = False

    ' 検索系は不要
    lstEmployees.Enabled = False
    btnSearch.Enabled = False

    ' --- ボタン制御 ---
    btnAdd.Enabled = True          ' 追加処理として使う
    btnUpdate.Enabled = False
    btnDelete.Enabled = False

'    'TEST↓------------------
'    Debug.Print "今のModeは" & CurrentMode
'    'TEST↑------------------

End Sub


'【SetMode_Update（検索前_更新モード）】
Private Sub SetMode_Update_before()

    CurrentMode = "Update_before"
   
    lblOperationGuide.Caption = ""
    lblOperationGuide.Caption = "更新対象を入力または選択し、検索ボタンを押してください"
    
    ' --- 検索前の状態 ---
    txtLastName.Enabled = True
    txtFirstName.Enabled = True
    TextEmpId.Enabled = True
    lstEmployees.Enabled = True
    btnSearch.Enabled = True

    ' 更新対象が確定するまで編集不可
    cmbDep.Enabled = False
    chkAdmin.Enabled = False

    ' --- ボタン制御 ---
    btnAdd.Enabled = False
    btnUpdate.Enabled = True      ' 更新処理として使う
    btnDelete.Enabled = False

'    'TEST↓------------------
'    Debug.Print "今のModeは" & CurrentMode
'    'TEST↑------------------

End Sub

'【SetMode_Update（検索後_更新モード）】
Private Sub SetMode_Update_after()

    CurrentMode = "Update"

    txtLastName.Enabled = True
    txtFirstName.Enabled = True
    TextEmpId.Enabled = False
    lstEmployees.Enabled = False
    btnSearch.Enabled = False

    cmbDep.Enabled = True
    chkAdmin.Enabled = True

    ' --- ボタン制御 ---
    btnAdd.Enabled = False
    btnUpdate.Enabled = True      ' 更新処理として使う
    btnDelete.Enabled = False
    
'    'TEST↓------------------
'    Debug.Print "今のModeは" & CurrentMode
'    'TEST↑------------------

End Sub

'【SetMode_Delete（検索前_削除モード）】
Private Sub SetMode_Delete_before()

    CurrentMode = "Delete_before"
    
    lblOperationGuide.Caption = ""
    lblOperationGuide.Caption = "削除する社員を入力か選択し、検索ボタンを押してください"
    
    ' --- 検索前の状態 ---
    txtLastName.Enabled = True
    txtFirstName.Enabled = True
    TextEmpId.Enabled = True
    lstEmployees.Enabled = True
    btnSearch.Enabled = True

    ' 削除は編集不要
    cmbDep.Enabled = False
    chkAdmin.Enabled = False

    ' --- ボタン制御 ---
    btnAdd.Enabled = False
    btnUpdate.Enabled = False
    btnDelete.Enabled = True      ' 削除処理として使う
    
'    'TEST↓------------------
'    Debug.Print "今のModeは" & CurrentMode
'    'TEST↑------------------

End Sub

'【SetMode_Delete（検索後_削除モード）】
Private Sub SetMode_Delete_after()

    CurrentMode = "Delete"

    ' --- 検索後の状態 ---
    txtLastName.Enabled = False
    txtFirstName.Enabled = False
    TextEmpId.Enabled = False
    lstEmployees.Enabled = False
    btnSearch.Enabled = False

    ' 削除は編集不要
    cmbDep.Enabled = False
    chkAdmin.Enabled = False

    ' --- ボタン制御 ---
    btnAdd.Enabled = False
    btnUpdate.Enabled = False
    btnDelete.Enabled = True      ' 削除処理として使う

'    'TEST↓------------------
'    Debug.Print "今のModeは" & CurrentMode
'    'TEST↑------------------

End Sub

'追加ボタン押下
Private Sub btnAdd_Click()
    
    ' --- モード未選択 → 追加モードに切替 ---
    If CurrentMode = "" Then
        Call SetMode_Add
        Exit Sub
    End If

    ' --- 追加モード中 → 追加処理を実行 ---
    If CurrentMode = "Add" Then
        Call Execute_Add
        Exit Sub
    End If

End Sub

'更新ボタン押下
Private Sub btnUpdate_Click()

    '検索前_更新モードへ切り替え
    If CurrentMode = "" Then
        Call SetMode_Update_before
        Exit Sub
    End If

    '検索前_更新ボタン押下エラー表示
    If CurrentMode = "Update_before" And Me.Tag = "" Then
        MsgBox "更新する従業員を検索または選択してください。", vbExclamation
        Exit Sub
    End If
    
    '更新処理を実行
    If CurrentMode = "Update" Then
        Call Execute_Update
        Exit Sub
    End If

End Sub

'削除ボタン押下
Private Sub btnDelete_Click()
    
    '検索前_削除モードへ切り替え
    If CurrentMode = "" Then
        Call SetMode_Delete_before
        Exit Sub
    End If

    '検索前_削除ボタン押下エラー表示
    If CurrentMode = "Delete_before" And Me.Tag = "" Then
        MsgBox "削除する従業員を検索または選択してください。", vbExclamation
        Exit Sub
    End If

    '削除処理を実行
    If CurrentMode = "Delete" Then
        Call Execute_Delete
        Exit Sub
    End If

End Sub

'検索ボタン押下
Private Sub btnSearch_Click()

    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long
    Dim searchName As String
    Dim searchID As String
    Dim found As Boolean
    Dim FullName As String
    Dim Parts As Variant
    
    Set ws = ThisWorkbook.Sheets(SHEET_EMP)
    lr = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)

    searchName = Trim(txtLastName.value) & " " & Trim(txtFirstName.value)
    searchID = Trim(TextEmpId.value)
    
    ' --- 入力チェック ---
    If searchID <> "" Then
    
        If Not IsNumeric(searchID) Then
            MsgBox "従業員IDは数値で入力してください。", vbExclamation
            Exit Sub
        End If
        
    Else
    
        If Trim(txtLastName.value) = "" And Trim(txtFirstName.value) = "" Then
            MsgBox "氏名または従業員IDを入力してください。", vbExclamation
            Exit Sub
        End If

        If Trim(txtLastName.value) = "" Then
            MsgBox "苗字も入力してください。", vbExclamation
            Exit Sub
        End If

        If Trim(txtFirstName.value) = "" Then
            MsgBox "名前も入力してください。", vbExclamation
            Exit Sub
        End If
        
    End If


    ' --- 検索処理 ---
    found = False
    For i = 2 To lr

        ' ID一致
        If searchID <> "" Then
            If ws.Cells(i, EMP_COL_EMPLOYEE_ID).value = CLng(searchID) Then
                found = True
            End If
        End If

        ' 氏名一致
        If searchName <> "" Then
            If ws.Cells(i, EMP_COL_NAME).value = searchName Then
                found = True
            End If
        End If

        ' 見つかった場合
        If found Then

            ' --- ListBox をクリア ---
            lstEmployees.Clear
            ' ListBox に表示
            lstEmployees.AddItem
            lstEmployees.list(lstEmployees.ListCount - 1, 0) = ws.Cells(i, EMP_COL_EMPLOYEE_ID).value
            lstEmployees.list(lstEmployees.ListCount - 1, 1) = ws.Cells(i, EMP_COL_NAME).value

            FullName = ws.Cells(i, EMP_COL_NAME).value
            Parts = Split(FullName, " ")
            ' フォームに反映
            txtLastName.value = Parts(0)
            txtFirstName.value = Parts(1)
            TextEmpId.value = ws.Cells(i, EMP_COL_EMPLOYEE_ID).value
            cmbDep.value = ws.Cells(i, EMP_COL_DEPARTMENT).value
            lblPass.Caption = "Password：" & ws.Cells(i, EMP_COL_PASSWORD).value
            chkAdmin.value = ws.Cells(i, EMP_COL_IS_ADMIN).value

            ' 行番号を保持（更新・削除で使用）
            Me.Tag = i

            Exit For
        End If

    Next i

'    'TEST↓------------------
'    Debug.Print "(Me.Tag)行番号:"; Me.Tag
'    'TEST↑------------------

    ' --- 見つからなかった場合 ---
    If Not found Then
        MsgBox "該当する従業員が見つかりません。", vbExclamation
        Me.Tag = ""
        
'        'TEST↓------------------
'        Debug.Print "(Me.Tag)行番号:"; Me.Tag
'        'TEST↑------------------
        
        Exit Sub
    End If


    lblOperationGuide.Caption = ""
    '検索前_更新モードなら
    If CurrentMode = "Update_before" Then
    lblOperationGuide.Caption = "内容を修正し、更新ボタンを押してください"
    End If
    
    '検索前_削除モードなら
    If CurrentMode = "Delete_before" Then
    lblOperationGuide.Caption = "内容を確認し、間違いなければ削除ボタンを押してください"
    End If

    '検索後_更新モードへ切り替え
    If CurrentMode = "Update_before" Then
        Call SetMode_Update_after
        Exit Sub
    End If

    '検索後_削除モードへ切り替え
    If CurrentMode = "Delete_before" Then
        Call SetMode_Delete_after
        Exit Sub
    End If

End Sub


'追加処理を実行
Private Sub Execute_Add()

    Dim ws As Worksheet
    Dim lr As Long
    Dim newID As Long
    Dim rawPass As String
    Dim hashedPass As String
    Dim FullName As String
    
    Set ws = ThisWorkbook.Sheets(SHEET_EMP)
    
    ' --- 入力チェック ---
    If Trim(txtLastName.value) = "" Then
        MsgBox "苗字を入力してください。", vbExclamation
        Exit Sub
    End If

    If Trim(txtFirstName.value) = "" Then
        MsgBox "名前を入力してください。", vbExclamation
        Exit Sub
    End If

    If Not IsValidName(txtLastName.value) Then
        MsgBox "苗字に使用できない文字が含まれています。" & vbCrLf & _
               "使用可能：ひらがな・カタカナ・漢字・英字・スペース（半角/全角）", vbExclamation
        Exit Sub
    End If

    If Not IsValidName(txtFirstName.value) Then
        MsgBox "名前に使用できない文字が含まれています。" & vbCrLf & _
               "使用可能：ひらがな・カタカナ・漢字・英字・スペース（半角/全角）", vbExclamation
        Exit Sub
    End If
    
    FullName = Trim(txtLastName.value) & " " & Trim(txtFirstName.value)
    
    If cmbDep.value = "" Then
        MsgBox "部門を選択してください。", vbExclamation
        Exit Sub
    End If

    ' --- 自動採番（EmployeeID） ---
    lr = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)
    If lr < 2 Then
        newID = 1
    Else
        newID = ws.Cells(lr, EMP_COL_EMPLOYEE_ID).value + 1
    End If

    ' --- 初期パスワード生成（8桁ランダム英数字） ---
    rawPass = GenerateRandomPassword(8)
    
    ' --- SHA256 ハッシュ化 ---
    hashedPass = SHA256_Hash(rawPass)

    ' --- Employees シートに追加 ---
    ws.Cells(lr + 1, EMP_COL_EMPLOYEE_ID).value = newID
    ws.Cells(lr + 1, EMP_COL_NAME).value = FullName
    ws.Cells(lr + 1, EMP_COL_DEPARTMENT).value = cmbDep.value
    ws.Cells(lr + 1, EMP_COL_PASSWORD).value = hashedPass
    ws.Cells(lr + 1, EMP_COL_IS_ADMIN).value = IIf(chkAdmin.value, 1, 0)

    ' --- 画面に初期パスワードを表示 ---
    lblPass.Caption = "Password：" & rawPass

    MsgBox "従業員を追加しました。" & vbCrLf & _
           "初期パスワード：" & rawPass, vbInformation

    ' --- 初期状態に戻す ---
    Call ResetToInitialMode
    
End Sub


'更新処理を実行
Private Sub Execute_Update()

    Dim ws As Worksheet
    Dim rowNum As Long

    Set ws = ThisWorkbook.Sheets(SHEET_EMP)

    ' --- 更新対象の行番号チェック ---
    If Me.Tag = "" Then
        MsgBox "更新する従業員を検索または選択してください。", vbExclamation
        Exit Sub
    End If

    rowNum = CLng(Me.Tag)

    ' --- 入力チェック ---
    If Trim(txtLastName.value) = "" Then
        MsgBox "苗字を入力してください。", vbExclamation
        Exit Sub
    End If

    If Trim(txtFirstName.value) = "" Then
        MsgBox "名前を入力してください。", vbExclamation
        Exit Sub
    End If

    If Not IsValidName(txtLastName.value) Then
        MsgBox "苗字に使用できない文字が含まれています。" & vbCrLf & _
               "使用可能：ひらがな・カタカナ・漢字・英字・スペース（半角/全角）", vbExclamation
        Exit Sub
    End If

    If Not IsValidName(txtFirstName.value) Then
        MsgBox "名前に使用できない文字が含まれています。" & vbCrLf & _
               "使用可能：ひらがな・カタカナ・漢字・英字・スペース（半角/全角）", vbExclamation
        Exit Sub
    End If

    If cmbDep.value = "" Then
        MsgBox "部門を選択してください。", vbExclamation
        Exit Sub
    End If

    ' --- Employees シートへ反映（パスワードは変更しない） ---
    ws.Cells(rowNum, EMP_COL_NAME).value = Trim(txtLastName.value) & " " & Trim(txtFirstName.value)
    ws.Cells(rowNum, EMP_COL_DEPARTMENT).value = cmbDep.value
    ws.Cells(rowNum, EMP_COL_IS_ADMIN).value = IIf(chkAdmin.value, 1, 0)

    MsgBox "従業員情報を更新しました。", vbInformation

    ' --- 初期状態に戻す ---
    Call ResetToInitialMode

End Sub



'削除処理を実行
Private Sub Execute_Delete()

    Dim ws As Worksheet
    Dim rowNum As Long
    Dim empName As String

    Set ws = ThisWorkbook.Sheets(SHEET_EMP)

    ' --- 削除対象の行番号チェック ---
    If Me.Tag = "" Then
        MsgBox "削除する従業員を検索または選択してください。", vbExclamation
        Exit Sub
    End If

    rowNum = CLng(Me.Tag)
    empName = ws.Cells(rowNum, EMP_COL_NAME).value

    ' --- 削除確認 ---
    If MsgBox("従業員「" & empName & "」を削除しますか？", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    ' --- Employees シートの該当行を削除 ---
    ws.Rows(rowNum).Delete

    MsgBox "従業員を削除しました。", vbInformation

    ' --- 初期状態に戻す ---
    Call ResetToInitialMode

End Sub



'戻るボタン
Private Sub btnBack_Click()
    Unload Me
    AdminMenuForm.Show
End Sub


'フォームを閉じる（開発モードでなければ完全に閉じる）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    HandleFormClose Me, Cancel, CloseMode
End Sub


