Attribute VB_Name = "TEST勤怠管理システム自動化"
Option Explicit

'
'' --- テスト用コード ↓---
''【SHA256 ハッシュ化関数】
'Public Function TEST_SHA256_Hash() As String
'    Dim shaObj As Object
'    Dim bytes() As Byte
'    Dim hash() As Byte
'    Dim i As Long
'    Dim result As String
'    Dim text As String
'
'    'text = "Abcdefg8"
'    text = "Fk1Hvk9f"
'
'    Set shaObj = CreateObject("System.Security.Cryptography.SHA256Managed")
'
'    bytes = StrConv(text, vbFromUnicode)
'    hash = shaObj.ComputeHash_2(bytes)
'
'    For i = LBound(hash) To UBound(hash)
'        result = result & LCase(Right("00" & Hex(hash(i)), 2))
'    Next i
'
'    Debug.Print text & "のハッシュ値は" & result
'Stop
'    TEST_SHA256_Hash = result
'End Function
'' --- ここまでテスト用 ↑---
