VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HistoryAllForm 
   ClientHeight    =   11436
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11784
   OleObjectBlob   =   "HistoryAllForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "HistoryAllForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

    
'Initialize
Private Sub UserForm_Initialize()

    lstListOfAllApplicationHistory.ColumnCount = 7
    lstListOfAllApplicationHistory.ColumnWidths = "68 pt;70 pt;70 pt;70 pt;156 pt;71 pt;35 pt"
    lstListOfAllApplicationHistory.Height = 280
    
    '色
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(236, 240, 241)
    btnBack.BackColor = RGB(52, 152, 219)
    btnBack.ForeColor = RGB(255, 255, 255)
    btnStaDate.BackColor = RGB(52, 152, 219)
    btnStaDate.ForeColor = RGB(255, 255, 255)
    btnEnDate.BackColor = RGB(52, 152, 219)
    btnEnDate.ForeColor = RGB(255, 255, 255)
    btnSearch.BackColor = RGB(52, 152, 219)
    btnSearch.ForeColor = RGB(255, 255, 255)
    btnReset.BackColor = RGB(52, 152, 219)
    btnReset.ForeColor = RGB(255, 255, 255)
    
    Call SetCmb
    Call InitForm
End Sub

'コンボボックス設定
Private Sub SetCmb()
    Dim ws As Worksheet
    Dim i As Long
    Dim lr As Long
    
    'シート存在チェック
    If Not SheetExists(SHEET_EMP) Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets(SHEET_EMP)
    lr = GetLastRow(ws, EMP_COL_EMPLOYEE_ID)
    
    'コンボボックスに社員名をセット
    cmbEmployeeName.Clear
    cmbEmployeeName.AddItem "全てを表示"
    For i = 2 To lr
        If Trim(ws.Cells(i, EMP_COL_NAME).value) <> "" Then
            cmbEmployeeName.AddItem ws.Cells(i, EMP_COL_NAME)
        End If
    Next i
    
    'プルダウンが画面に収まりきらない場合はスクロールバーを表示
    cmbEmployeeName.ListRows = 15
   
    '手入力禁止、リストからの選択のみ許可
    cmbEmployeeName.Style = fmStyleDropDownList
    
    'コンボボックスに申請ステータスをセット
    cmbStatus.Clear
    cmbStatus.AddItem "全てを表示"
    cmbStatus.AddItem "申請中"
    cmbStatus.AddItem "承認済"
    cmbStatus.AddItem "却下済"
    '手入力禁止、リストからの選択のみ許可
    cmbStatus.Style = fmStyleDropDownList
            
End Sub
    
'Initialize/リセットボタンで使用
Private Sub InitForm()

    lstListOfAllApplicationHistory.Clear
    
    'TextBox初期化
    If txtStaDate.value <> "" Then
        txtStaDate.value = ""
    End If

    If txtEnDate.value <> "" Then
        txtEnDate.value = ""
    End If
    
    'コンボボックスの初期値を"全てを表示"に設定
    cmbEmployeeName.ListIndex = 0
    cmbStatus.ListIndex = 0

    Call GetAllAttendanceSheetsName
    Call SetDateForAllUsersAndAllStatuses

End Sub


'休暇開始日のボタンを押下
Private Sub btnStaDate_Click()

    If txtStaDate.value <> "" Then
        txtStaDate.value = ""
    End If
    
    CalendarForm.CalendarMode = "AllApplications"
    Set CalendarForm.TargetTextbox = Me.txtStaDate
    CalendarForm.Show
End Sub


'休暇終了日のボタンを押下
Private Sub btnEnDate_Click()

    If txtEnDate.value <> "" Then
        txtEnDate.value = ""
    End If
    
    CalendarForm.CalendarMode = "AllApplications"
    Set CalendarForm.TargetTextbox = Me.txtEnDate
    CalendarForm.Show
End Sub



'全ユーザー・全ステータスの日付を設定する
Private Sub SetDateForAllUsersAndAllStatuses()
    Dim stDate() As Date
    Dim enDate() As Date
    Dim TargetID() As Long
    Dim Types() As String
    Dim Remarks() As String

    Call GetHolidayFromAllUsersAndAllStatuses(stDate, enDate, TargetID, Types, Remarks)
    
    '空配列チェック
    If (Not Not stDate) = 0 Then
        MsgBox "休暇申請はありません"
        Exit Sub
    End If
    
    Call Sort_Attendance
    Call LoadAllApplications(stDate(), enDate(), TargetID(), Types(), Remarks())

End Sub


'引数を受け取りListBox表示する
Private Function LoadAllApplications(START_Date() As Date, END_Date() As Date, Linked_EmpID() As Long, Leave_Type() As String, Remarks() As String)

    Dim wsAtt As Worksheet
    Dim lrAtt As Long
    Dim i As Long
    Dim k As Long
    Dim targetStart As Date
    Dim targetEnd  As Date
    Dim TargetID As Long
    Dim targetName As String
    Dim targetType As String
    Dim targetRemarks As String
    Dim sheetName As String

    lstListOfAllApplicationHistory.Clear
          
    For k = LBound(START_Date) To UBound(START_Date)
        targetStart = START_Date(k)
        targetEnd = END_Date(k)
        TargetID = Linked_EmpID(k)
        targetType = Leave_Type(k)
        targetRemarks = Remarks(k)

        If targetStart = 0 Then GoTo ContinueLoop
        
        'targetStart の年月からシート名を作成
        sheetName = "Attendance_" & year(targetStart) & "_" & Format(month(targetStart), "00")

        'シートが存在しない場合はスキップ
        If Not SheetExists(sheetName) Then GoTo ContinueLoop
        
            Set wsAtt = ThisWorkbook.Sheets(sheetName)
            lrAtt = GetLastRow(wsAtt, ATT_COL_EMPLOYEE_ID)
                       
            For i = 2 To lrAtt
                If TargetID = wsAtt.Cells(i, ATT_COL_EMPLOYEE_ID).value And targetStart = CDate(wsAtt.Cells(i, ATT_COL_DATE).value) _
                And targetType = wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value And targetRemarks = wsAtt.Cells(i, ATT_COL_REMARKS).value Then
                
                    targetName = GetEmployeeNameByID(TargetID)
                
                    lstListOfAllApplicationHistory.AddItem
                    lstListOfAllApplicationHistory.list(lstListOfAllApplicationHistory.ListCount - 1, 0) = targetName                                     '申請者名
                    lstListOfAllApplicationHistory.list(lstListOfAllApplicationHistory.ListCount - 1, 1) = wsAtt.Cells(i, ATT_COL_APPLICATION_DATE).value '申請日
                    lstListOfAllApplicationHistory.list(lstListOfAllApplicationHistory.ListCount - 1, 2) = targetStart                                    '休暇開始日
                    lstListOfAllApplicationHistory.list(lstListOfAllApplicationHistory.ListCount - 1, 3) = targetEnd                                      '休暇終了日
                    lstListOfAllApplicationHistory.list(lstListOfAllApplicationHistory.ListCount - 1, 4) = wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value       '種別
                    lstListOfAllApplicationHistory.list(lstListOfAllApplicationHistory.ListCount - 1, 5) = wsAtt.Cells(i, ATT_COL_REMARKS).value          '備考
                    lstListOfAllApplicationHistory.list(lstListOfAllApplicationHistory.ListCount - 1, 6) = wsAtt.Cells(i, ATT_COL_STATUS).value           '承認ステータス
                    
                    Exit For
                End If
            Next i
ContinueLoop:
    Next k
End Function


'検索ボタン
Private Sub btnSearch_Click()
    Dim stDate() As Date
    Dim enDate() As Date
    Dim TargetID() As Long
    Dim Types() As String
    Dim Remarks() As String
    Dim selName As String
    Dim selID As Long
    Dim selStatus As String
    Dim selStaDate  As Date
    Dim selEnDate  As Date

    '入力チェック
    If txtStaDate.value = "" And txtEnDate.value <> "" _
    Or txtStaDate.value <> "" And txtEnDate.value = "" Then
        MsgBox "日付の指定は、開始日と終了日の両方を入れてください。"
        Exit Sub
    End If
    
    '前回の検索結果を削除
    lstListOfAllApplicationHistory.Clear
    
    '選択値を受け取り変数へ格納
    If cmbEmployeeName.value <> "全てを表示" Then
        selName = cmbEmployeeName.value
              
        '社員名を社員IDに変換
        selID = GetEmployeeIDByName(selName) '辞書を用いた関数を使用(modUtil)
    End If
    
    If cmbStatus.value <> "全てを表示" Then
        selStatus = cmbStatus.value
    End If
    
    If txtStaDate.value <> "" And txtEnDate.value <> "" Then
        selStaDate = Trim(txtStaDate.value)
        selEnDate = Trim(txtEnDate.value)
    End If
    
    Call SearchHolidayRequests(selName, selID, selStatus, selStaDate, selEnDate, stDate, enDate, TargetID, Types, Remarks)
    
    '空配列チェック
    If (Not Not stDate) = 0 Then
        lstListOfAllApplicationHistory.Clear
        MsgBox "該当する休暇申請はありません"
        Exit Sub
    End If
    
    Call LoadAllApplications(stDate(), enDate(), TargetID(), Types(), Remarks())

End Sub


'リセットボタン
Private Sub btnReset_Click()
    Call InitForm
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


