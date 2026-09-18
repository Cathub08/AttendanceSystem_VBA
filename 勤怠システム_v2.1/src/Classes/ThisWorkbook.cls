VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



'Excel ファイルを開いた瞬間に LoginForm が表示される
Private Sub Workbook_Open()
    'Excel感を消す
    RemoveExcelLook
    LoginForm.Show
End Sub


'Excelを閉じる前の処理
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RestoreExcelLook
End Sub



