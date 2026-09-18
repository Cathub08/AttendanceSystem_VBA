VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HistoryForm 
   ClientHeight    =   9132.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10788
   OleObjectBlob   =   "HistoryForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "HistoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    lstApplication.Clear
    lstApplication.ColumnCount = 6
    lstApplication.ColumnWidths = "70 pt;70 pt;70 pt;160 pt;90 pt;20 pt"

    '色
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(236, 240, 241)
    btnCancel.BackColor = RGB(52, 152, 219)
    btnCancel.ForeColor = RGB(255, 255, 255)
    btnBack.BackColor = RGB(52, 152, 219)
    btnBack.ForeColor = RGB(255, 255, 255)
       
End Sub

Private Sub UserForm_Activate()
    Call SetTheDate
End Sub


'休暇開始日・終了日を取得する
Private Sub SetTheDate()
    
    Dim stDate() As Date
    Dim enDate() As Date
    
    Call GetConsecutiveHolidays(stDate, enDate)
      
    '空配列チェック
    If (Not Not stDate) = 0 Then
        MsgBox "休暇申請はありません"
        Exit Sub
    End If
   
    Call Sort_Attendance
    Call LoadApplication(stDate(), enDate())

End Sub


'引数を受け取りListBox表示する
Private Function LoadApplication(START_Date() As Date, END_Date() As Date)

    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long
    Dim k As Long
    Dim empID As Long
    Dim Counts As Long
    Dim targetStart As Date
    Dim targetEnd  As Date
    Dim sheetName As String
    
    empID = GetLoggedInID() 'ログイン中の従業員IDをempIDに入れる
    lstApplication.Clear
                                           
    For k = LBound(START_Date) To UBound(START_Date)
        targetStart = START_Date(k)
        targetEnd = END_Date(k)
        
        If targetStart = 0 Then GoTo ContinueLoop

        'targetStart の年月からシート名を作成
        sheetName = "Attendance_" & year(targetStart) & "_" & Format(month(targetStart), "00")
        
        'シートが存在しない場合はスキップ
        If Not SheetExists(sheetName) Then GoTo ContinueLoop
        
            Set ws = ThisWorkbook.Sheets(sheetName)
            lr = GetLastRow(ws, ATT_COL_EMPLOYEE_ID)
                       
            For i = 2 To lr
                If empID = ws.Cells(i, ATT_COL_EMPLOYEE_ID).value And ws.Cells(i, ATT_COL_LEAVE_TYPE).value <> "" _
                And targetStart = CDate(ws.Cells(i, ATT_COL_DATE).value) Then
                    lstApplication.AddItem
                    lstApplication.list(lstApplication.ListCount - 1, 0) = ws.Cells(i, ATT_COL_APPLICATION_DATE).value '申請日
                    lstApplication.list(lstApplication.ListCount - 1, 1) = targetStart                                 '休暇開始日
                    lstApplication.list(lstApplication.ListCount - 1, 2) = targetEnd                                   '休暇終了日
                    lstApplication.list(lstApplication.ListCount - 1, 3) = ws.Cells(i, ATT_COL_LEAVE_TYPE).value       '種別
                    lstApplication.list(lstApplication.ListCount - 1, 4) = ws.Cells(i, ATT_COL_REMARKS).value          '備考
                    lstApplication.list(lstApplication.ListCount - 1, 5) = ws.Cells(i, ATT_COL_STATUS).value           '承認ステータス
                    
                    Exit For
                End If
            Next i
            
ContinueLoop:
    Next k
End Function

 

'取り消しボタン
Private Sub btnCancel_Click()
    Dim selID As Long
    Dim selStaDate  As Date
    Dim selEnDate  As Date
    Dim selLT As String
    Dim selRemark As String

    If lstApplication.ListIndex = -1 Then
        MsgBox "申請を選択してください。", vbExclamation
        Exit Sub
    End If

    selID = GetLoggedInID()
    selStaDate = CDate(lstApplication.list(lstApplication.ListIndex, 1)) '休暇開始日
    selEnDate = CDate(lstApplication.list(lstApplication.ListIndex, 2))  '休暇終了日
    selLT = lstApplication.list(lstApplication.ListIndex, 3)             '種別
    selRemark = lstApplication.list(lstApplication.ListIndex, 4)         '備考

    Dim res  As result
    res = CancelLeaveRequest(selID, selStaDate, selEnDate, selLT, selRemark)
        Select Case res
            Case Success
            
                Call Sort_Attendance
                
                Call SetTheDate
                
                MsgBox "取消が完了しました", vbInformation
                
            Case InvalidInput
                MsgBox "申請中のみ取消可能です。", vbExclamation
                
            Case Other
                MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
            End Select
End Sub




'閉じるボタン
Private Sub btnBack_Click()
    Unload Me
    LeaveForm.Show
End Sub


'フォームを閉じる（開発モードでなければ完全に閉じる）
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    HandleFormClose Me, Cancel, CloseMode
End Sub

