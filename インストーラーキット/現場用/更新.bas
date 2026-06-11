Attribute VB_Name = "更新"
Option Explicit

'================================================================================
' [更新] 更新（「更新」ボタンの本体 → マスタから工事番号一覧を取り込む）
'--------------------------------------------------------------------------------
' 【何のファイル？】
'   マスタの工事番号一覧シートを、現場用xlsm内のローカルコピー
'   シートに上書きする処理。AutoFilter で自部署のみ表示にする。
'
' 【しくみ】
'   - コピー先シート名は管理マスタの「工事番号一覧シート」項目から取得
'   - B列（連番NO）は非表示にする（工事番号で一意管理のため不要）
'================================================================================


' IS_TEST_MODE / TEST_FILE_PATH / SHEET_* / CELL_* は Config モジュールで一元管理。

'================================================================================
' 「工事番号一覧」と「依頼履歴」を両方更新するマクロ
'================================================================================
Sub UpdateAllSheets()
    Dim originalScreenUpdating As Boolean
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim wsOriginal As Worksheet
    Set wsOriginal = ActiveSheet  ' 元のアクティブシートを記憶（最後に戻すため）

    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerUpdateAll

    Call UpdateKoujiBangoListSheet(False)
    Call UpdateIraiRirekiSheet(False)

    ' 元のシートに戻す（AutoFilter適用で勝手にシート移動するのを防ぐ）
    On Error Resume Next
    If Not wsOriginal Is Nothing Then wsOriginal.Activate
    On Error GoTo 0

    MsgBox "「工事番号一覧」と「依頼履歴」シートが正常に更新されました。", vbInformation, "更新完了"

FinalizeUpdateAll:
    ' 例外時も元のシートに戻す
    On Error Resume Next
    If Not wsOriginal Is Nothing Then wsOriginal.Activate
    On Error GoTo 0
    Application.ScreenUpdating = originalScreenUpdating
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub

ErrorHandlerUpdateAll:
    MsgBox "一括更新中にエラーが発生しました: " & Err.Description, vbCritical, "更新エラー"
    Resume FinalizeUpdateAll
End Sub

'================================================================================
' 「工事番号一覧」シートを最新の情報に更新するマクロ
'================================================================================
Sub UpdateKoujiBangoListSheet(Optional ByVal ShowMessage As Boolean = True)
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsMaster As Worksheet
    Dim wsDest As Worksheet
    Dim targetFilePath As String
    Dim destSheetName As String
    Dim lastRowSource As Long
    Dim copyRange As Range

    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalScreenUpdating As Boolean

    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerUpdateList

    targetFilePath = GetTargetFilePath()

    If Dir(targetFilePath) = "" Then
        MsgBox "対象のファイルが見つかりません。" & vbCrLf & targetFilePath, vbCritical, "ファイルエラー"
        GoTo FinalizeUpdateList
    End If

    On Error Resume Next
    Set wbTarget = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo ErrorHandlerUpdateList

    If wbTarget Is Nothing Then
        MsgBox "対象のExcelファイルを開けませんでした。" & vbCrLf & _
               "他のユーザーが使用中である可能性があります。", vbCritical
        GoTo FinalizeUpdateList
    End If

    If Not SheetExists(wbTarget, SHEET_KOUJI_LIST) Or Not SheetExists(wbTarget, SHEET_KANRI_MASTER) Then
        MsgBox "外部ファイルに必要なシート「" & SHEET_KOUJI_LIST & "」または「" & SHEET_KANRI_MASTER & "」が見つかりません。", vbCritical, "シートエラー"
        GoTo FinalizeUpdateList
    End If
    Set wsSource = wbTarget.Sheets(SHEET_KOUJI_LIST)
    Set wsMaster = wbTarget.Sheets(SHEET_KANRI_MASTER)

    destSheetName = GetSheetNameFromMaster(ThisWorkbook, "工事番号一覧シート", SHEET_KOUJI_LIST)
    If destSheetName = "" Then
        MsgBox "「管理マスタ」シートの「工事番号一覧シート」項目に値が設定されていません。", vbExclamation
        GoTo FinalizeUpdateList
    End If

    If Not SheetExists(ThisWorkbook, destSheetName) Then
        MsgBox "このファイルにコピー先のシート「" & destSheetName & "」が見つかりませんでした。", vbExclamation
        GoTo FinalizeUpdateList
    End If
    Set wsDest = ThisWorkbook.Sheets(destSheetName)

    Call SafeUnprotect(wsDest)
    wsDest.Range("A3:X" & wsDest.Rows.count).Clear

    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    If lastRowSource >= 5 Then
        Set copyRange = wsSource.Range("A5:X" & lastRowSource)
        copyRange.Copy Destination:=wsDest.Range("A3")
    End If

    ' 自部門の行のみ表示するAutoFilterを適用（D列=工事番号）
    Call ApplyDeptFilter(wsDest)

    ' ApplyDeptFilter 内で SafeProtectData により保護されているので、
    ' 一旦 SafeUnprotect してから Hidden を設定する。
    Call SafeUnprotect(wsDest)

    ' B列「NO」(連番) を非表示にする（2026/6/1～ 工事番号だけで管理する仕様変更による）
    '   従来は担当者番号と組み合わせて連番管理していたが、工事番号自体が一意の通番となったため
    '   B列の連番表示は不要との運用方針。
    On Error Resume Next
    wsDest.Columns("B").Hidden = True
    ' V・W列も非表示にする【2026/6/11 お客様要望】
    wsDest.Columns("V").Hidden = True
    wsDest.Columns("W").Hidden = True
    On Error GoTo 0

    If ShowMessage Then
        MsgBox "「" & destSheetName & "」シートが正常に更新されました。", vbInformation, "更新完了"
    End If
    Call SafeProtectData(wsDest)

FinalizeUpdateList:
    Application.CutCopyMode = False
    If Not wsMaster Is Nothing Then Set wsMaster = Nothing
    If Not wsSource Is Nothing Then Set wsSource = Nothing
    If Not wsDest Is Nothing Then Set wsDest = Nothing
    If Not wbTarget Is Nothing Then wbTarget.Close SaveChanges:=False

    Application.ScreenUpdating = originalScreenUpdating
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub

ErrorHandlerUpdateList:
    If Not wsDest Is Nothing Then Call SafeProtectData(wsDest)  ' エラー時も保護を復元
    MsgBox "「" & SHEET_KOUJI_LIST & "」シートの更新中にエラーが発生しました: " & Err.Description, vbCritical, "更新エラー"
    Resume FinalizeUpdateList
End Sub

'================================================================================
' 「依頼履歴」シートを最新の情報に更新するマクロ
'================================================================================
Sub UpdateIraiRirekiSheet(Optional ByVal ShowMessage As Boolean = True)
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim targetFilePath As String
    Dim lastRowSource As Long
    Dim copyRange As Range

    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalScreenUpdating As Boolean

    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerUpdateRireki

    targetFilePath = GetTargetFilePath()

    If Dir(targetFilePath) = "" Then
        MsgBox "対象のファイルが見つかりません。" & vbCrLf & targetFilePath, vbCritical, "ファイルエラー"
        GoTo FinalizeUpdateRireki
    End If

    On Error Resume Next
    Set wbTarget = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo ErrorHandlerUpdateRireki

    If wbTarget Is Nothing Then
        MsgBox "対象のExcelファイルを開けませんでした。" & vbCrLf & _
               "他のユーザーが使用中である可能性があります。", vbCritical
        GoTo FinalizeUpdateRireki
    End If

    If Not SheetExists(wbTarget, SHEET_IRAI_RIREKI) Then
        MsgBox "外部ファイルに必要なシート「" & SHEET_IRAI_RIREKI & "」が見つかりません。", vbCritical, "シートエラー"
        GoTo FinalizeUpdateRireki
    End If
    Set wsSource = wbTarget.Sheets(SHEET_IRAI_RIREKI)

    If Not SheetExists(ThisWorkbook, SHEET_IRAI_RIREKI) Then
        MsgBox "このファイルにコピー先のシート「" & SHEET_IRAI_RIREKI & "」が見つかりませんでした。", vbExclamation
        GoTo FinalizeUpdateRireki
    End If
    Set wsDest = ThisWorkbook.Sheets(SHEET_IRAI_RIREKI)

    Call SafeUnprotect(wsDest)
    wsDest.Range("A3:W" & wsDest.Rows.count).Clear

    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    If lastRowSource >= 2 Then
        Set copyRange = wsSource.Range("A2:W" & lastRowSource)
        copyRange.Copy Destination:=wsDest.Range("A3")
    End If

    ' 自部門の行のみ表示するAutoFilterを適用（D列=工事番号）
    Call ApplyDeptFilter(wsDest, 5)  ' E列=工事番号

    If ShowMessage Then
        MsgBox "「" & SHEET_IRAI_RIREKI & "」シートが正常に更新されました。", vbInformation, "更新完了"
    End If
    Call SafeProtectData(wsDest)

FinalizeUpdateRireki:
    Application.CutCopyMode = False
    If Not wsSource Is Nothing Then Set wsSource = Nothing
    If Not wsDest Is Nothing Then Set wsDest = Nothing
    If Not wbTarget Is Nothing Then wbTarget.Close SaveChanges:=False

    Application.ScreenUpdating = originalScreenUpdating
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub

ErrorHandlerUpdateRireki:
    If Not wsDest Is Nothing Then Call SafeProtectData(wsDest)  ' エラー時も保護を復元
    MsgBox "「" & SHEET_IRAI_RIREKI & "」シートの更新中にエラーが発生しました: " & Err.Description, vbCritical, "更新エラー"
    Resume FinalizeUpdateRireki
End Sub

'================================================================================
' 工事番号一覧シートに自部門のAutoFilterを適用
'   D列(工事番号)が "<dept>-*" で始まる行のみ表示する
'   将来全社展開で部署混在になっても、各ファイルが自部門だけ見える
'================================================================================
Public Sub ApplyDeptFilter(ByVal ws As Worksheet, Optional ByVal koujiBangouColIndex As Long = 4)
    ' ws: 対象シート
    ' koujiBangouColIndex: 範囲先頭からの「工事番号」列のインデックス
    '   工事番号一覧 (A～X): D列=4 (既定)
    '   依頼履歴   (A～W): E列=5
    If ws Is Nothing Then Exit Sub
    Dim deptCode As String: deptCode = GetMyDeptCode()
    If deptCode = "" Then Exit Sub

    On Error Resume Next
    Call SafeUnprotect(ws)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 3 Then GoTo ProtectAndExit  ' データ最低1行（R3）必要

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A2:X" & lastRow).AutoFilter Field:=koujiBangouColIndex, _
        Criteria1:=deptCode & "-*"

ProtectAndExit:
    Call SafeProtectData(ws)
    On Error GoTo 0
End Sub
