Attribute VB_Name = "modConst"
Option Explicit

'======================
' modConst：Enum・定数
'======================

'Enum(列挙型定数)　result(結果)
Public Enum result

    Success = 0             '成功
    
    InvalidInput = 1        'ユーザーの入力ミス
    
    FileOpen = 10           'ユーザーが対処できるシステム異常(10-100)
    
    NoRecordsForToday = 11  '本日の出勤記録なし(MsgBox:"本日の出勤記録がありません。")
    
    NoData = 12             'データが存在しない(MsgBox:"今月の勤怠シートの内容がありません。月次レポートは出力されません。")
    
    Other = 100             'ユーザーが対処できないシステム異常(MsgBox:システム内部のエラーにより処理を完了できませんでした。)

End Enum



'Sheet name
Public Const SHEET_MAIN As String = "Main"
Public Const SHEET_EMP As String = "Employees"
Public Const SHEET_DEP As String = "Departments"
Public Const SHEET_LEAVE As String = "Leave Type"
Public Const SHEET_HOLIDAY As String = "HolidayMaster"
Public Const SHEET_M_REPO_TEMPLATE As String = "MonthlyReport_Template"
Public Const SHEET_ATT_TEMPLATE As String = "Attendance_Template"
Public Const SHEET_ATT_PREFIX As String = "Attendance_"


'所定労働時間
Public Const STD_WORK_HOURS As Double = 8


'Attendance sheet column index
Public Const ATT_COL_EMPLOYEE_ID As Long = 1       ' 社員番号
Public Const ATT_COL_DATE As Long = 2              ' 日付
Public Const ATT_COL_CLOCK_IN As Long = 3          ' 出勤
Public Const ATT_COL_CLOCK_OUT As Long = 4         ' 退勤
Public Const ATT_COL_BREAK_TIME As Long = 5        ' 休憩時間
Public Const ATT_COL_WORK_HOURS As Long = 6        ' 実働時間
Public Const ATT_COL_OVERTIME As Long = 7          ' 残業
Public Const ATT_COL_LEAVE_TYPE As Long = 8        ' 休暇種別
Public Const ATT_COL_REMARKS As Long = 9           ' 備考
Public Const ATT_COL_APPLICATION_DATE As Long = 10 ' 申請日
Public Const ATT_COL_STATUS As Long = 11           ' 申請状態


'Employees sheet column index
Public Const EMP_COL_EMPLOYEE_ID As Long = 1   ' 社員番号
Public Const EMP_COL_NAME As Long = 2          ' 社員名
Public Const EMP_COL_DEPARTMENT As Long = 3    ' 部署名
Public Const EMP_COL_PASSWORD As Long = 4      ' パスワード
Public Const EMP_COL_IS_ADMIN As Long = 5      ' 管理者


'Departments sheet column index
Public Const DEP_COL_DEPARTMENT As Long = 1    ' 部署名


'LeaveType sheet column index
Public Const LT_COL_LEAVE_TYPE As Long = 1     ' 休暇種別


'HolidayMaster sheet column index
Public Const HM_COL_YEAR As Long = 1           ' 年
Public Const HM_COL_DATE As Long = 2           ' 日付
Public Const HM_COL_HOLIDAY_NAME As Long = 3   ' 名称


' Monthly　Report Template Column　index
Public Const MRepo_COL_EMPLOYEE_ID As Long = 1           ' 社員ID
Public Const MRepo_COL_EMPLOYEE_NAME As Long = 2         ' 氏名
Public Const MRepo_COL_NUMBER_OF_WORKING_DAY As Long = 3 ' 出勤日数
Public Const MRepo_COL_NUMBER_OF_DAYS_ABSENT As Long = 4 ' 欠勤日数
Public Const MRepo_COL_PAID_LEAVE_DAYS As Long = 5       ' 有給取得日数
Public Const MRepo_COL_TOTAL_WORKING_HOURS As Long = 6   ' 勤務時間合計
Public Const MRepo_COL_TOTAL_OVERTIME_HOURS As Long = 7  ' 残業時間合計
Public Const MRepo_COL_TOTAL_HOURS As Long = 8           ' 総労働時間


'開発中→True,開発終了→False
Public Const Development_Mode As Boolean = True

