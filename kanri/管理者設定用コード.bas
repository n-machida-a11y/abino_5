Attribute VB_Name = "管理者設定用コード"
Option Explicit

'================================================================================
' 管理者用 共通設定モジュール
'================================================================================

' ===== テストモード =====
Public Const IS_TEST_MODE As Boolean = True

Public Const TEST_MASTER_PATH As String = "Z:\Users\n-machida\Desktop\工事番号管理表.xlsm"
Public Const PROD_MASTER_PATH As String = "Z:\全社共有\建築事業部\30_事務\工事番号管理\工事番号管理表.xlsm"

' ===== シート保護パスワード =====
Public Const SHEET_PASSWORD As String = "3555"

' ===== スナップショットシートのプレフィックス =====
' 最新取得時点のマスタ状態をこの名前で記録する（反映時の三者比較用）
Public Const SNAPSHOT_PREFIX As String = "_snap_"

'================================================================================
' 各シートの同期設定
'   Array(シート名, キー列文字, データ開始行, モード)
'   モード: "merge"     = キーベースの三者マージ
'           "overwrite" = 全上書き（キー概念が無いシート用）
'
' ※ "2025.6～個人番号" "2025.6～採番ルール" はVBAから参照されていない
'   参考ドキュメントなので同期対象から除外
'================================================================================
Public Function GetSheetSyncConfig() As Variant
    Dim configs(4) As Variant
    configs(0) = Array("工事番号一覧",        "D",  4, "merge")  ' 工事番号がキー
    configs(1) = Array("依頼履歴",           "A",  3, "merge")  ' 依頼NOがキー
    configs(2) = Array("管理マスタ",         "B",  2, "merge")  ' 担当者番号がキー
    configs(3) = Array("その他マスタ",        "A",  2, "merge")  ' 提出先名がキー
    configs(4) = Array("依頼書セル設定",      "A", 10, "merge")  ' 項目名がキー
    GetSheetSyncConfig = configs
End Function

'================================================================================
' マスタファイルのパスを返す共通関数
'================================================================================
Public Function GetMasterPath() As String
    If IS_TEST_MODE Then
        GetMasterPath = TEST_MASTER_PATH
    Else
        GetMasterPath = PROD_MASTER_PATH
    End If
End Function

'================================================================================
' 指定ブック内に特定名のシートが存在するか
'================================================================================
Public Function SheetExistsIn(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExistsIn = Not (ws Is Nothing)
End Function

'================================================================================
' シート保護ヘルパー
'================================================================================
Public Sub SafeUnprotect(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0
End Sub

Public Sub SafeProtect(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=SHEET_PASSWORD, _
               UserInterfaceOnly:=False, _
               AllowFiltering:=True, _
               AllowSorting:=True, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True
    On Error GoTo 0
End Sub

Public Sub ClearAllFilters(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
End Sub
