VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ApproveForm 
   ClientHeight    =   9432.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11784
   OleObjectBlob   =   "ApproveForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ApproveForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'Initialize
Private Sub UserForm_Initialize()

    lstListOfAllApplications.Clear
    lstListOfAllApplications.ColumnCount = 7
    lstListOfAllApplications.ColumnWidths = "68 pt;70 pt;70 pt;70 pt;160 pt;72 pt;50 pt"
    lstListOfAllApplications.Height = 300

    '色
    lblTitle.BackColor = RGB(44, 62, 80)
    lblTitle.ForeColor = RGB(255, 255, 255)
    Me.BackColor = RGB(236, 240, 241)
    btnBack.BackColor = RGB(52, 152, 219)
    btnBack.ForeColor = RGB(255, 255, 255)
    btnApproval.BackColor = RGB(52, 152, 219)
    btnApproval.ForeColor = RGB(255, 255, 255)
    btnRejected.BackColor = RGB(52, 152, 219)
    btnRejected.ForeColor = RGB(255, 255, 255)
    
End Sub


Private Sub UserForm_Activate()
    Call SetDateForAllUsers
End Sub

'すべてのユーザーの日付を設定する
Private Sub SetDateForAllUsers()
    Dim stDate() As Date
    Dim enDate() As Date
    Dim TargetID() As Long
    
    Call GetConsecutiveHolidaysOfAllUsers(stDate, enDate, TargetID)
      
    '空配列チェック
    If (Not Not stDate) = 0 Then
        MsgBox "休暇申請はありません"
        Exit Sub
    End If

    Call Sort_Attendance
    Call LoadAllApplications(stDate(), enDate(), TargetID())

End Sub


'引数を受け取りListBox表示する
Private Function LoadAllApplications(START_Date() As Date, END_Date() As Date, Linked_EmpID() As Long)

    Dim wsAtt As Worksheet
    Dim lrAtt As Long
    Dim i As Long
    Dim k As Long
    Dim targetStart As Date
    Dim targetEnd  As Date
    Dim TargetID As Long
    Dim targetName As String
    Dim sheetName As String
    
    lstListOfAllApplications.Clear
                    
    For k = LBound(START_Date) To UBound(START_Date)
        targetStart = START_Date(k)
        targetEnd = END_Date(k)
        TargetID = Linked_EmpID(k)

        If targetStart = 0 Then GoTo ContinueLoop
        
        'targetStart の年月からシート名を作成
        sheetName = "Attendance_" & year(targetStart) & "_" & Format(month(targetStart), "00")

        'シートが存在しない場合はスキップ
        If Not SheetExists(sheetName) Then GoTo ContinueLoop
    
            Set wsAtt = ThisWorkbook.Sheets(sheetName)
            lrAtt = GetLastRow(wsAtt, ATT_COL_EMPLOYEE_ID)
                       
            For i = 2 To lrAtt
                If wsAtt.Cells(i, ATT_COL_EMPLOYEE_ID).value = TargetID _
                And wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value <> "" _
                And targetStart = CDate(wsAtt.Cells(i, ATT_COL_DATE).value) Then

                    targetName = GetEmployeeNameByID(TargetID)
                
                    lstListOfAllApplications.AddItem
                    lstListOfAllApplications.list(lstListOfAllApplications.ListCount - 1, 0) = targetName                                     '申請者名
                    lstListOfAllApplications.list(lstListOfAllApplications.ListCount - 1, 1) = wsAtt.Cells(i, ATT_COL_APPLICATION_DATE).value '申請日
                    lstListOfAllApplications.list(lstListOfAllApplications.ListCount - 1, 2) = targetStart                                    '休暇開始日
                    lstListOfAllApplications.list(lstListOfAllApplications.ListCount - 1, 3) = targetEnd                                      '休暇終了日
                    lstListOfAllApplications.list(lstListOfAllApplications.ListCount - 1, 4) = wsAtt.Cells(i, ATT_COL_LEAVE_TYPE).value       '種別
                    lstListOfAllApplications.list(lstListOfAllApplications.ListCount - 1, 5) = wsAtt.Cells(i, ATT_COL_REMARKS).value          '備考
                    lstListOfAllApplications.list(lstListOfAllApplications.ListCount - 1, 6) = wsAtt.Cells(i, ATT_COL_STATUS).value           '承認ステータス
                    
                    Exit For
                End If
            Next i
ContinueLoop:
    Next k
End Function

 

'承認ボタン
Private Sub btnApproval_Click()

    Dim selName As String
    Dim selStaDate  As Date
    Dim selEnDate  As Date
    Dim selType As String
    Dim selRemark As String

    '申請を未選択でボタンを押下した場合
    If lstListOfAllApplications.ListIndex = -1 Then
        MsgBox "申請を選択してください。"
        Exit Sub
    End If
    
    selName = lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 0)           '申請者名
    selStaDate = CDate(lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 2)) '休暇開始日
    selEnDate = CDate(lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 3))  '休暇終了日
    selType = lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 4)           '種別
    selRemark = lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 5)         '備考

    Dim res As result
    res = ApprovalRequest(selName, selStaDate, selEnDate, selType, selRemark)
        Select Case res
            Case Success
            
                Call SetDateForAllUsers
                MsgBox "承認が完了しました", vbInformation
                
            Case InvalidInput
                MsgBox "該当月の勤怠シートが存在しないため、承認できません。", vbExclamation
                
            Case Other
                MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
        End Select
                
End Sub


'却下ボタン
Private Sub btnRejected_Click()

    Dim selName As String
    Dim selStaDate  As Date
    Dim selEnDate  As Date
    Dim selType As String
    Dim selRemark As String

    '申請を未選択でボタンを押下した場合
    If lstListOfAllApplications.ListIndex = -1 Then
        MsgBox "申請を選択してください。"
        Exit Sub
    End If
    
    selName = lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 0)           '申請者名
    selStaDate = CDate(lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 2)) '休暇開始日
    selEnDate = CDate(lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 3))  '休暇終了日
    selType = lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 4)           '種別
    selRemark = lstListOfAllApplications.list(lstListOfAllApplications.ListIndex, 5)         '備考

    Dim res As result
    res = RejectionRequest(selName, selStaDate, selEnDate, selType, selRemark)
        Select Case res
            Case Success
                Call SetDateForAllUsers
                MsgBox "却下が完了しました", vbInformation
                
            Case InvalidInput
                MsgBox "該当月の勤怠シートが存在しないため、却下できません。", vbExclamation
            
            Case Other
                MsgBox "システム内部のエラーにより処理を完了できませんでした｡", vbCritical
        End Select

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


