Attribute VB_Name = "modScreenControl"
Option Explicit

'modScreenControl

Public Sub RemoveExcelLook()
    'Excel感を消す
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
End Sub

Public Sub RestoreExcelLook()
    'Excel感を戻す
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayHeadings = True
End Sub


'終了処理
Public Sub Exit_Operation()
 
    'Excel感を戻す
    RestoreExcelLook

    'Mainシートを表示
    If SheetExists(SHEET_MAIN) Then
        ThisWorkbook.Sheets(SHEET_MAIN).Activate
    End If

    Logout

End Sub


'フォームを閉じる処理
Public Sub HandleFormClose(frm As Object, Cancel As Integer, CloseMode As Integer)

    If CloseMode <> vbFormControlMenu Then Exit Sub

    If Development_Mode Then
        Exit_Operation
        Cancel = False   '開発モードはForm閉じてOK
        Exit Sub
    End If

    ' 本番モード
    Cancel = True       '(先に)Formが閉じない
    Exit_Operation
    ThisWorkbook.Close SaveChanges:=True 'ここでExcel自体が終了する(一緒にFormも終了)

End Sub


'シートの視認性をセットする
Public Sub SetSheetVisibility()

    Dim ws As Worksheet
    Dim item As Variant

    '全シートを非表示（Mainだけ例外）
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> SHEET_MAIN Then
            ws.Visible = xlSheetHidden
        End If
    Next ws

    'テンプレートは常に非表示
    ThisWorkbook.Sheets(SHEET_M_REPO_TEMPLATE).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(SHEET_ATT_TEMPLATE).Visible = xlSheetVeryHidden

    'Attendance の月次シートだけ動的に表示
    For Each ws In ThisWorkbook.Worksheets
        If ws.name Like SHEET_ATT_PREFIX & "*" _
           And ws.name <> SHEET_ATT_TEMPLATE Then
    
            ws.Visible = xlSheetVisible
        End If
    Next ws

End Sub

