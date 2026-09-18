Attribute VB_Name = "modHolidaySheet"
Option Explicit

'===========================================
' modHolidaySheet：休暇マスタ管理（LTシート）
'===========================================

'休暇種別マスタ（Leave Typeシートをアクティブにする）
Public Function Act_LT_sheet() As result
    On Error GoTo ErrHandler
    
    '初期値
    Act_LT_sheet = Other
        
    If Not SheetExists(SHEET_LEAVE) Then
            Act_LT_sheet = InvalidInput
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_LEAVE)
    ws.Activate
    Act_LT_sheet = Success
    Exit Function
    
ErrHandler:
    Act_LT_sheet = Other '想定外の異常
End Function



