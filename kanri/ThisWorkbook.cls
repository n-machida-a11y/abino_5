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
' [ファイル] ThisWorkbook（管理者入力用xlsm）
'--------------------------------------------------------------------------------
' 【何のファイル？】
'   Excelファイルを開いたとき・閉じたときに自動で動く処理を書く場所。
'   ここでは「シート保護を再設定」「不要なシート表示を整理」などの
'   開閉時メンテをまとめている。
'
' 【さわるとき】
'   ほぼ触る必要なし。シート保護パスワードが変わったときだけ
'   「管理者設定用コード」の SHEET_PASSWORD を変えればOK。
'================================================================================


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
