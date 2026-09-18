Attribute VB_Name = "modAuth"
Option Explicit

'===============
' modAuth：認証
'===============

'LoginFormで値を格納される
Public LoggedInID As Long

'ログイン処理
Public Function Login(empID As Long, pwd As String) As result
On Error GoTo ErrHandler

    Dim hashedInput As String
    Dim ws As Worksheet
    Dim lr As Long
             
    '初期値
    Login = Other
    
    'シート存在チェック
    If Not SheetExists(SHEET_EMP) Then Exit Function
    
    Set ws = ThisWorkbook.Sheets(SHEET_EMP)
    hashedInput = SHA256_Hash(pwd) ' ← 入力された平文をハッシュ化
    lr = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)

    Dim i As Long
    For i = 2 To lr
        If ws.Cells(i, EMP_COL_EMPLOYEE_ID).value = empID Then
            If Trim(ws.Cells(i, EMP_COL_PASSWORD).value) = hashedInput Then
                Dim isAdmin As Boolean
                isAdmin = CBool(Trim(ws.Cells(i, EMP_COL_IS_ADMIN).value)) 'Trimで余分なスペースによるエラーを回避
                Call SetLoginInfo(empID, ws.Cells(i, EMP_COL_NAME).value, isAdmin)
                Login = Success
                Exit Function
            End If
        End If
    Next i
    
    Login = InvalidInput
    Exit Function
    
ErrHandler:
    Login = Other
End Function


'ログイン情報の表示をクリア
Sub Logout()
    ClearLoginInfo
End Sub



'modLogin＞LoggedInID→LoginFormで値を格納されるグローバル変数
Public Function GetLoggedInID() As Long
    GetLoggedInID = LoggedInID
End Function


'ログイン中の社員名を取得
Public Function GetLoggedInName() As String

    Dim ws As Worksheet
    Dim lrEmp As Long
    Dim empID As Variant
    empID = GetLoggedInID()

    If empID = "" Then
        GetLoggedInName = ""
        Exit Function
    End If

    'シート存在チェック
    If Not SheetExists(SHEET_EMP) Then Exit Function
        
    Set ws = ThisWorkbook.Sheets(SHEET_EMP)
    lrEmp = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)
    
    Dim i As Long
    For i = 2 To lrEmp
        If ws.Cells(i, EMP_COL_EMPLOYEE_ID).value = empID Then
            GetLoggedInName = ws.Cells(i, EMP_COL_NAME).value
            Exit Function
        End If
    Next i

    GetLoggedInName = ""

End Function



'ログイン情報をMainシートへ表示
Public Sub SetLoginInfo(empID As Variant, empName As String, isAdmin As Boolean)
    
    'シート存在チェック
    If Not SheetExists(SHEET_MAIN) Then Exit Sub
    
    With ThisWorkbook.Sheets(SHEET_MAIN)
        .Range("B2").value = empID
        .Range("B3").value = empName
        .Range("B4").value = IIf(isAdmin, "管理者", "一般")
    End With
End Sub



'Mainシートのログイン情報をクリア
Public Sub ClearLoginInfo()

    'シート存在チェック
    If Not SheetExists(SHEET_MAIN) Then Exit Sub
    
    With ThisWorkbook.Sheets(SHEET_MAIN)
        .Range("B2:B4").ClearContents
    End With
End Sub


'引数のEmpIDが管理者か判定
Function isAdmin(empID As Long) As result
On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lrEmp As Long

    '初期値
    isAdmin = Other

    'シート存在チェック
    If Not SheetExists(SHEET_EMP) Then Exit Function '想定内の異常
   
    Set ws = ThisWorkbook.Sheets(SHEET_EMP)

    Dim i As Long
    lrEmp = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)
    For i = 2 To lrEmp

        If ws.Cells(i, EMP_COL_EMPLOYEE_ID).value = empID Then
            'E列（5列目）が管理者フラグ（1=管理者）
            If CLng(ws.Cells(i, EMP_COL_IS_ADMIN).value) = 1 Then
                isAdmin = Success
                Exit Function
            End If
            
        End If

    Next i
    
    isAdmin = InvalidInput
    Exit Function
    
ErrHandler:
    isAdmin = Other '想定外の異常
End Function


'isAdminをBoolean型で返却
Public Function IsAdminByID(empID As Long) As Boolean
    IsAdminByID = (isAdmin(empID) = Success)
End Function


