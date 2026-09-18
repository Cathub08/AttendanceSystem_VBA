Attribute VB_Name = "modSecurity"
Option Explicit

'===============================
' modSecurity：パスワード＋暗号
'===============================

'パスワード更新処理　ChangePasswordFormで使用。
Public Function UpdatePassword(empID As Long, pwd As String, NewPwd As String) As result
    On Error GoTo ErrHandler
    
    Dim hashedInput As String
    Dim NewHashedInput As String
    Dim ws As Worksheet
    Dim lr As Long
    
    '初期値
    UpdatePassword = Other
    
    'シート存在チェック
    If Not SheetExists(SHEET_EMP) Then Exit Function
        
    Set ws = ThisWorkbook.Sheets(SHEET_EMP)
    lr = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)
    hashedInput = SHA256_Hash(pwd)       ' ← 旧PWをハッシュ化
    NewHashedInput = SHA256_Hash(NewPwd) ' ← 新PWをハッシュ化

    Dim i As Long
    For i = 2 To lr
        If Trim(ws.Cells(i, EMP_COL_EMPLOYEE_ID).value) = empID Then
            If Trim(ws.Cells(i, EMP_COL_PASSWORD).value) = hashedInput Then
                    ws.Cells(i, EMP_COL_PASSWORD).value = NewHashedInput
                    UpdatePassword = Success
                    Exit Function
            End If
        End If
    Next i

    UpdatePassword = InvalidInput
    Exit Function
    
ErrHandler:
    UpdatePassword = Other '想定外の異常
End Function



'入力チェック(英数字8桁)　ChangePasswordFormで使用。
Public Function IsValidPwd(NewPwd As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

    ' 英数字混在8桁チェック
    reg.Pattern = "^(?=.*[A-Za-z])(?=.*[0-9])[A-Za-z0-9]{8}$"
    reg.IgnoreCase = True
    reg.Global = True
    
    IsValidPwd = reg.TEST(NewPwd)
End Function



'【SHA256 ハッシュ化関数】
Public Function SHA256_Hash(text As String) As String
    Dim shaObj As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim i As Long
    Dim result As String

    Set shaObj = CreateObject("System.Security.Cryptography.SHA256Managed")

    bytes = StrConv(text, vbFromUnicode)
    hash = shaObj.ComputeHash_2(bytes)

    For i = LBound(hash) To UBound(hash)
        result = result & LCase(Right("00" & Hex(hash(i)), 2))
    Next i

    SHA256_Hash = result
End Function


'【ランダムパスワード生成関数】
Public Function GenerateRandomPassword(length As Long) As String

    Dim chars As String
    Dim i As Long
    Dim result As String

    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"

    Randomize

    For i = 1 To length
        result = result & Mid(chars, Int(Rnd() * Len(chars)) + 1, 1)
    Next i

    GenerateRandomPassword = result

End Function



