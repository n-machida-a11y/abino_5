Attribute VB_Name = "Config"
Option Explicit

'================================================================================
' モジュール: Config（【経理用】請求書作成.xlsm 用）
' 概要: このブックの全モジュールで共通に使う定数・設定をここだけで管理する。
'       設定変更はこのファイルを修正するだけでOK。
'================================================================================

' ===== テストモード =====
' ★★★ True=テスト / False=本番  ← ここ1箇所を変えるだけで全モジュールに反映 ★★★
Public Const IS_TEST_MODE As Boolean = False
Public Const TEST_MASTER_FILE_PATH As String = "Z:\Users\n-machida\Desktop\工事番号管理表.xlsm"

' ===== シート名 =====
Public Const SHEET_IRAI_RIREKI As String = "依頼履歴"
Public Const SHEET_IRAI_SEARCH As String = "依頼検索"

' ===== 依頼検索シートの「マスターファイルパス」格納セル =====
' （依頼履歴シートのG1セルにパスが書かれている）
Private Const MASTER_PATH_CELL_KEIRI As String = "G1"

'================================================================================
' マスターファイルのパスを返す共通関数
' 以前は Keiri_シート更新 内に定義されていた GetMasterPath() をここに集約。
' Keiri_保存 から Application.Run("GetMasterPath") で呼んでいたものも、
' 同一ブック内なので GetMasterPath() と直接呼べるようになる。
'================================================================================
Public Function GetMasterPath() As String
    If IS_TEST_MODE Then
        GetMasterPath = TEST_MASTER_FILE_PATH
    Else
        GetMasterPath = Trim(CStr(ThisWorkbook.Sheets(SHEET_IRAI_RIREKI).Range(MASTER_PATH_CELL_KEIRI).Value))
    End If
End Function
