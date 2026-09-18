Attribute VB_Name = "modEmployee"
Option Explicit

'========================
' modEmployee：社員データ
'========================

'社員IDの存在チェック　*LoginFormで使用。
Public Function EmployeeExists(empID As Long) As result
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long
    
    '初期値
    EmployeeExists = Other
        
    'シート存在チェック
    If Not SheetExists(SHEET_EMP) Then Exit Function
   
    Set ws = ThisWorkbook.Sheets(SHEET_EMP)
    lr = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)

    For i = 2 To lr
        If empID = Trim(ws.Cells(i, EMP_COL_EMPLOYEE_ID).value) Then
            EmployeeExists = Success
            Exit Function
        End If
    Next i
    
    EmployeeExists = InvalidInput
    Exit Function
    
ErrHandler:
    EmployeeExists = Other '想定外の異常
End Function



'"社員ID→社員名"変換の辞書を作成し、指定したIDの社員名を取得
Public Function GetEmployeeNameByID(TargetID As Long) As String
  
    Dim wsEmp As Worksheet
    Dim lrEmp As Long
    Dim dictName As Object
    Dim i As Long
        
    'シート存在チェック
    If Not SheetExists(SHEET_EMP) Then Exit Function

    Set wsEmp = ThisWorkbook.Sheets(SHEET_EMP)
    lrEmp = GetLastRow(wsEmp, EMP_COL_EMPLOYEE_ID)
    
      '社員ID→社員名の辞書を作成（全社員分）
    Set dictName = CreateObject("Scripting.Dictionary")
                    
    For i = 2 To lrEmp
        dictName(wsEmp.Cells(i, EMP_COL_EMPLOYEE_ID).value) = wsEmp.Cells(i, EMP_COL_NAME).value
    Next i
    
    'TargetID から 該当する社員名を取得
    If dictName.Exists(TargetID) Then
        GetEmployeeNameByID = dictName(TargetID)
    Else
        GetEmployeeNameByID = ""
    End If
    
End Function



'"社員名→社員ID"変換の辞書を作成し、指定した社員名のIDを取得
Public Function GetEmployeeIDByName(selectedName As String) As Long
    Dim wsEmp  As Worksheet
    Dim lrEmp As Long
    Dim dictID As Object
    Dim i As Long

    'シート存在チェック
    If Not SheetExists(SHEET_EMP) Then Exit Function
    
    Set wsEmp = ThisWorkbook.Sheets(SHEET_EMP)
    lrEmp = GetLastRow(wsEmp, EMP_COL_EMPLOYEE_ID)

    '社員名→社員ID の辞書を作成（全社員分）
    Set dictID = CreateObject("Scripting.Dictionary")

    For i = 2 To lrEmp
        dictID(wsEmp.Cells(i, EMP_COL_NAME).value) = wsEmp.Cells(i, EMP_COL_EMPLOYEE_ID).value
    Next i

    'selectedName から該当するIDを取得
    If dictID.Exists(selectedName) Then
        GetEmployeeIDByName = dictID(selectedName)
    Else
        GetEmployeeIDByName = 0
    End If

End Function



'【氏名バリデーション(検証)関数】
Public Function IsValidName(name As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

    ' 長さチェック
    If Len(name) < 1 Or Len(name) > 30 Then
        IsValidName = False
        Exit Function
    End If

    ' 許可する文字だけで構成されているかチェック
    ' ひらがな：\u3040-\u309F
    ' カタカナ：\u30A0-\u30FF
    ' 漢字：\u4E00-\u9FFF
    '々
    ' 英字：A-Za-z
    ' スペース：半角・全角
    reg.Pattern = "^[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF々 A-Za-z　]+$"
    reg.IgnoreCase = True
    reg.Global = True
    
    IsValidName = reg.TEST(name)
End Function

