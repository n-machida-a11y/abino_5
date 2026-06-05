Attribute VB_Name = "管理者設定用コード"
Option Explicit

'================================================================================
' [設定] 管理者設定用コード（環境ごとに値を変える場所）
'--------------------------------------------------------------------------------
' 【何のファイル？】
'   管理者用xlsm の「環境依存の設定値」をひとまとめにした場所。
'   テスト⇔本番の切替、マスタファイルのパス、部署コード等。
'
' 【お客様環境セットアップ時にチェックする箇所】
'   1. IS_TEST_MODE       … 本番なら必ず False
'   2. PROD_MASTER_PATH   … 本番のマスタxlsm のフルパス
'   3. 「操作」シート B1  … 自部署の2桁コード（建築=03/土木=05）
'   4. 「操作」シート A1  … 本番モード時のマスタパス（空欄なら PROD_MASTER_PATH）
'
' 【触らなくていい箇所】
'   - SHEET_PASSWORD（保護パスワード）
'   - SNAPSHOT_PREFIX（内部処理用の接頭辞）
'   - 各種 Function（パス取得・部署コード取得などのヘルパー）
'================================================================================


'================================================================================
' 管理者用 共通設定モジュール
'================================================================================

' ===== テストモード =====
Public Const IS_TEST_MODE As Boolean = False

Public Const TEST_MASTER_PATH As String = "Z:\Users\n-machida\Desktop\工事番号管理表.xlsm"
Public Const PROD_MASTER_PATH As String = "Z:\全社共有\建築事業部\00_現場関連\管理者用\工事番号管理表.xlsm"

' ===== シート保護パスワード =====
Public Const SHEET_PASSWORD As String = "3555"

' ===== 自部門コード =====
' この管理者ファイルが管理する部門の2桁コード（建築事業部=03 など）
' 【2026/5/22 改修】
'   旧: 定数 MY_DEPT_CODE を VBE で書き換える運用
'   新: 「操作」シートの B1 セルから GetMyDeptCode() で取得する運用
'        → VBE 書換不要、シート上で設定変更できる
'   既存コードとの互換性のため定数は「フォールバック既定値」として残す。
Public Const MY_DEPT_CODE_DEFAULT As String = "03"

' 互換用：既存コードから参照される可能性がある定数（GetMyDeptCode を使うのが推奨）
Public Const MY_DEPT_CODE As String = "03"

' ===== スナップショットシートのプレフィックス =====
' 最新取得時点のマスタ状態をこの名前で記録する（反映時の三者比較用）
' 【2026/5/22 位置修正】
' VBAルール「モジュール定数は Function/Sub の前」のため、ここに配置。
Public Const SNAPSHOT_PREFIX As String = "_snap_"

'================================================================================
' GetMyDeptCode: 「操作」シートの B1 セルから部門コードを取得
'================================================================================
' 配布時の運用：
'   建築事業部用: 「操作」シート B1 セルに "03" を入力
'   土木事業部用: 「操作」シート B1 セルに "05" を入力
'   空欄なら MY_DEPT_CODE_DEFAULT ("03") にフォールバック
'================================================================================
Public Function GetMyDeptCode() As String
    On Error Resume Next
    Dim v As String
    v = Trim(CStr(ThisWorkbook.Sheets("操作").Range("B1").Value))
    On Error GoTo 0
    If v = "" Then v = MY_DEPT_CODE_DEFAULT
    GetMyDeptCode = v
End Function

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
    ' Array(シート名, キー列, データ開始行, モード, 部門フィルタ列文字)
    ' 部門フィルタ列が "" でない場合、その列の値が "MY_DEPT_CODE-*" で始まる行のみ
    ' 同期対象にする。部署別管理者ファイルで自部門のレコードだけ扱うため。
    Dim configs(5) As Variant
    configs(0) = Array("工事番号一覧",        "D",  5, "merge",     "D")  ' 工事番号がキー、D列で部門フィルタ
    configs(1) = Array("依頼履歴",           "A",  2, "merge",     "E")  ' 依頼NOがキー、E列(工事番号)で部門フィルタ
    configs(2) = Array("管理マスタ",         "A",  2, "overwrite", "")    ' 全部署共通・key/value 形式のため全上書き
    configs(3) = Array("その他マスタ",        "A",  2, "merge",     "")    ' 全部署共通
    configs(4) = Array("依頼書セル設定",      "A", 10, "merge",     "")    ' 全部署共通
    ' 【2026/6/5 変更】担当者マスタは merge → overwrite に変更。
    '   複数部門を担当する人（企画開発）は同じ名前で部門ごとに複数行登録するが、
    '   merge は担当者名をキーにするため同名行が1行に潰れてしまう実害が発生した。
    '   全行コピー(overwrite)なら複数行登録がそのまま保たれる。
    '   ※overwrite は「最後に反映した人が勝つ」ため、担当者マスタの編集は
    '     マスタ側で直接行うか、編集担当を一人に決める運用とすること。
    configs(5) = Array("担当者マスタ",        "A",  2, "overwrite", "")    ' 全部署共通
    GetSheetSyncConfig = configs
End Function

'================================================================================
' マスタファイルのパスを返す共通関数
'--------------------------------------------------------------------------------
' 【2026/5/26 改修】本番モード時に「操作」シート A1 セルでパスを指定可能にした。
'   優先順位:
'     1. IS_TEST_MODE = True  → コードの TEST_MASTER_PATH を使う（A1は無視）
'                               ※ 開発・社内テスト用
'     2. IS_TEST_MODE = False → 「操作」シート A1 セルのパスを使う
'                               A1 が空欄なら PROD_MASTER_PATH にフォールバック
'
' 【お客様がテスト用マスタで動作確認する手順】
'   ・本番マスタをコピーして例: デスクトップに「工事番号管理表_テスト.xlsm」を作る
'   ・「操作」シート A1 セルにそのコピーのフルパスを貼り付ける
'   ・「データ更新」を実行してテスト
'   ・確認できたら A1 を空欄に戻す（本番マスタに自動で戻る）
'================================================================================
Public Function GetMasterPath() As String
    If IS_TEST_MODE Then
        ' テストモード: コードの TEST_MASTER_PATH を使う（A1セルは無視）
        GetMasterPath = TEST_MASTER_PATH
    Else
        ' 本番モード: 「操作」シート A1 セルのパスを使う
        '             A1 が空欄なら PROD_MASTER_PATH にフォールバック
        Dim cellPath As String
        On Error Resume Next
        cellPath = Trim(CStr(ThisWorkbook.Sheets("操作").Range("A1").Value))
        On Error GoTo 0
        If cellPath <> "" Then
            GetMasterPath = cellPath
        Else
            GetMasterPath = PROD_MASTER_PATH
        End If
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
