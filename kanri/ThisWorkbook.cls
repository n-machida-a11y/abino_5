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

'================================================================================
' ThisWorkbook イベント（管理者用）
'================================================================================

' ブックを開いた時、操作シートを自動で表示する
' （マクロ_最新取得 / マクロ_マスタへ反映 ボタンへの動線を明確にするため）
Private Sub Workbook_Open()
    On Error Resume Next
    ThisWorkbook.Sheets("操作").Activate
    On Error GoTo 0
End Sub
