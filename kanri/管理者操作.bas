Attribute VB_Name = "管理者操作"
Option Explicit

'================================================================================
' 管理者用 メイン操作モジュール（差分マージ版）
'
' ■ 動作
'   ・最新取得時: マスタ内容をこのブックに転写し、同時に隠しスナップショットも作成
'   ・反映時    : 3者比較（ローカル／マスタ／スナップショット）で差分マージ
'                 → 管理者の編集・追加・削除を反映しつつ、他者が追加した行は保持
'================================================================================


'================================================================================
' 【ボタン1】マスタから最新取得
'================================================================================
Public Sub マクロ_最新取得()
    Dim masterPath As String: masterPath = GetMasterPath()
    If Dir(masterPath) = "" Then
        MsgBox "マスタファイルにアクセスできません。" & vbCrLf & _
               "パス: " & masterPath, vbCritical, "最新取得エラー"
        Exit Sub
    End If

    If MsgBox("マスタから最新データを取得します。" & vbCrLf & _
              "このブックの対象シートとスナップショットは上書きされます。" & vbCrLf & vbCrLf & _
              "よろしいですか？", vbYesNo + vbQuestion, "最新取得") = vbNo Then Exit Sub

    Dim wbMaster As Workbook
    Dim configs As Variant, cfg As Variant
    Dim logText As String
    Dim successCount As Long, skipCount As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Cleanup

    Set wbMaster = Workbooks.Open(fileName:=masterPath, ReadOnly:=True, UpdateLinks:=0)

    configs = GetSheetSyncConfig()
    Dim i As Long
    For i = LBound(configs) To UBound(configs)
        cfg = configs(i)
        Dim sheetName As String: sheetName = CStr(cfg(0))

        If Not SheetExistsIn(wbMaster, sheetName) Then
            logText = logText & " - " & sheetName & " (マスタに無し、スキップ)" & vbCrLf
            skipCount = skipCount + 1
            GoTo NextSheet
        End If
        If Not SheetExistsIn(ThisWorkbook, sheetName) Then
            logText = logText & " - " & sheetName & " (ローカルに無し、スキップ)" & vbCrLf
            skipCount = skipCount + 1
            GoTo NextSheet
        End If

        Dim srcSh As Worksheet, dstSh As Worksheet
        Set srcSh = wbMaster.Sheets(sheetName)
        Set dstSh = ThisWorkbook.Sheets(sheetName)

        Call SafeUnprotect(dstSh)
        Call ClearAllFilters(dstSh)

        ' ローカル可視シートへコピー
        dstSh.Cells.Clear
        If srcSh.UsedRange.Cells.CountLarge > 0 Then
            srcSh.UsedRange.Copy
            dstSh.Range(srcSh.UsedRange.Address).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            dstSh.Range(srcSh.UsedRange.Address).PasteSpecial Paste:=xlPasteColumnWidths
        End If
        Application.CutCopyMode = False

        ' スナップショットシートへもコピー（反映時の3者比較用）
        Call WriteSnapshotFromSheet(srcSh, sheetName)

        successCount = successCount + 1
        logText = logText & " - " & sheetName & " (取得成功)" & vbCrLf
NextSheet:
    Next i

    wbMaster.Close SaveChanges:=False
    Set wbMaster = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "最新取得が完了しました。" & vbCrLf & vbCrLf & _
           "成功: " & successCount & " シート" & vbCrLf & _
           "スキップ: " & skipCount & " シート" & vbCrLf & vbCrLf & _
           "詳細:" & vbCrLf & logText, vbInformation, "最新取得完了"
    Exit Sub

Cleanup:
    If Not wbMaster Is Nothing Then wbMaster.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "最新取得中にエラーが発生しました: " & vbCrLf & Err.Description, vbCritical
End Sub


'================================================================================
' 【ボタン2】マスタへ反映（差分マージ）
'================================================================================
Public Sub マクロ_マスタへ反映()
    Dim masterPath As String: masterPath = GetMasterPath()
    If Dir(masterPath) = "" Then
        MsgBox "マスタファイルにアクセスできません。" & vbCrLf & _
               "パス: " & masterPath, vbCritical, "反映エラー"
        Exit Sub
    End If

    ' スナップショット存在確認
    Dim configs As Variant: configs = GetSheetSyncConfig()
    Dim i As Long, missingSnapshots As String
    For i = LBound(configs) To UBound(configs)
        Dim sn As String: sn = SNAPSHOT_PREFIX & CStr(configs(i)(0))
        If Not SheetExistsIn(ThisWorkbook, sn) Then
            If CStr(configs(i)(3)) = "merge" Then
                missingSnapshots = missingSnapshots & " - " & CStr(configs(i)(0)) & vbCrLf
            End If
        End If
    Next i
    If missingSnapshots <> "" Then
        If MsgBox("以下のシートでスナップショットが未作成です:" & vbCrLf & missingSnapshots & vbCrLf & _
                  "差分マージができないため、他者の追加が失われる可能性があります。" & vbCrLf & vbCrLf & _
                  "先に「最新取得」を実行することを強く推奨します。" & vbCrLf & vbCrLf & _
                  "このまま進めますか？（全上書きフォールバック）", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "警告") = vbNo Then Exit Sub
    End If

    If MsgBox("ローカルの変更をマスタへ反映します。" & vbCrLf & vbCrLf & _
              "差分マージ方式:" & vbCrLf & _
              " ・管理者の編集・追加・削除を反映" & vbCrLf & _
              " ・取得後に他者が追加した行は保持" & vbCrLf & vbCrLf & _
              "続行しますか？", vbYesNo + vbQuestion, "マスタへ反映") = vbNo Then Exit Sub

    ' 他Excelで開かれていないかチェック
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase(wb.FullName) = LCase(masterPath) Then
            MsgBox "マスタファイルが別のExcelウィンドウで開かれています。閉じてから再実行してください。", vbCritical
            Exit Sub
        End If
    Next wb

    Dim wbMaster As Workbook
    Dim logText As String
    Dim successCount As Long, skipCount As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Cleanup

    Set wbMaster = Workbooks.Open(fileName:=masterPath, ReadOnly:=False, UpdateLinks:=0)
    If wbMaster.ReadOnly Then
        MsgBox "マスタファイルを書き込み可能で開けませんでした。他のユーザーが編集中の可能性があります。", vbCritical
        GoTo Cleanup
    End If

    For i = LBound(configs) To UBound(configs)
        Dim cfg As Variant: cfg = configs(i)
        Dim sheetName As String: sheetName = CStr(cfg(0))
        Dim keyCol As String: keyCol = CStr(cfg(1))
        Dim dataStartRow As Long: dataStartRow = CLng(cfg(2))
        Dim mode As String: mode = CStr(cfg(3))

        If Not SheetExistsIn(wbMaster, sheetName) Then
            logText = logText & " - " & sheetName & " (マスタに無し、スキップ)" & vbCrLf
            skipCount = skipCount + 1
            GoTo NextSheet
        End If
        If Not SheetExistsIn(ThisWorkbook, sheetName) Then
            logText = logText & " - " & sheetName & " (ローカルに無し、スキップ)" & vbCrLf
            skipCount = skipCount + 1
            GoTo NextSheet
        End If

        Dim ret As String
        If mode = "overwrite" Or keyCol = "" Then
            ret = SyncSheetOverwrite(ThisWorkbook.Sheets(sheetName), wbMaster.Sheets(sheetName))
        Else
            Dim snapSheet As Worksheet
            If SheetExistsIn(ThisWorkbook, SNAPSHOT_PREFIX & sheetName) Then
                Set snapSheet = ThisWorkbook.Sheets(SNAPSHOT_PREFIX & sheetName)
            Else
                Set snapSheet = Nothing
            End If
            ret = SyncSheetMerge( _
                    ThisWorkbook.Sheets(sheetName), _
                    wbMaster.Sheets(sheetName), _
                    snapSheet, _
                    keyCol, dataStartRow)
        End If

        logText = logText & " - " & sheetName & ": " & ret & vbCrLf
        successCount = successCount + 1
NextSheet:
    Next i

    wbMaster.Save
    wbMaster.Close SaveChanges:=False
    Set wbMaster = Nothing

    ' 反映成功したらスナップショットも新しいマスタ状態で更新しておく
    Call SilentRefreshSnapshots

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "マスタへの反映が完了しました。" & vbCrLf & vbCrLf & _
           "成功: " & successCount & " シート" & vbCrLf & _
           "スキップ: " & skipCount & " シート" & vbCrLf & vbCrLf & _
           "詳細:" & vbCrLf & logText, vbInformation, "反映完了"
    Exit Sub

Cleanup:
    If Not wbMaster Is Nothing Then wbMaster.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "反映中にエラーが発生しました: " & vbCrLf & Err.Description, vbCritical
End Sub


'================================================================================
' シート単位: 全上書きモード
'================================================================================
Public Function SyncSheetOverwrite(ByVal src As Worksheet, ByVal dst As Worksheet) As String
    Call SafeUnprotect(dst)
    Call ClearAllFilters(dst)
    dst.Cells.Clear
    If src.UsedRange.Cells.CountLarge > 0 Then
        src.UsedRange.Copy
        dst.Range(src.UsedRange.Address).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        dst.Range(src.UsedRange.Address).PasteSpecial Paste:=xlPasteColumnWidths
    End If
    Application.CutCopyMode = False
    SyncSheetOverwrite = "全上書き完了"
End Function


'================================================================================
' シート単位: 差分マージモード（キーベース3者比較）
'   local   : 管理者のローカルシート（編集済み）
'   master  : マスタシート（現在の本物）
'   snap    : スナップショット（最新取得時点のマスタ）
'   keyCol  : キー列文字 (例: "D")
'   dataRow : データ開始行
'================================================================================
Public Function SyncSheetMerge(ByVal localSh As Worksheet, ByVal masterSh As Worksheet, ByVal snapSh As Worksheet, ByVal keyCol As String, ByVal dataRow As Long) As String

    ' 各キー→行番号 のマップを作る
    Dim dictLocal As Object, dictMaster As Object, dictSnap As Object
    Set dictLocal = BuildKeyMap(localSh, keyCol, dataRow)
    Set dictMaster = BuildKeyMap(masterSh, keyCol, dataRow)
    If snapSh Is Nothing Then
        Set dictSnap = CreateObject("Scripting.Dictionary")  ' 空（他者追加と区別不可）
    Else
        Set dictSnap = BuildKeyMap(snapSh, keyCol, dataRow)
    End If

    Call SafeUnprotect(masterSh)

    Dim lastCol As Long
    lastCol = localSh.UsedRange.Columns.Count
    If masterSh.UsedRange.Columns.Count > lastCol Then lastCol = masterSh.UsedRange.Columns.Count
    If Not snapSh Is Nothing Then
        If snapSh.UsedRange.Columns.Count > lastCol Then lastCol = snapSh.UsedRange.Columns.Count
    End If
    If lastCol < 1 Then lastCol = 1

    Dim updated As Long, added As Long, deleted As Long, preserved As Long
    Dim k As Variant

    ' 1. localSh にある = 編集か追加
    For Each k In dictLocal.Keys
        Dim srcRow As Long: srcRow = dictLocal(k)
        If dictMaster.Exists(k) Then
            ' 編集: masterの該当行をlocalで上書き
            Dim dstRow1 As Long: dstRow1 = dictMaster(k)
            Call CopyRow(localSh, srcRow, masterSh, dstRow1, lastCol)
            updated = updated + 1
        Else
            ' 追加: masterに新規行として追記
            Dim newRow As Long
            newRow = GetLastDataRow(masterSh, keyCol) + 1
            If newRow < dataRow Then newRow = dataRow
            Call CopyRow(localSh, srcRow, masterSh, newRow, lastCol)
            added = added + 1
        End If
    Next k

    ' 2. masterSh にあるが localSh に無い = 削除 or 他者追加
    Dim rowsToDelete As Collection: Set rowsToDelete = New Collection
    For Each k In dictMaster.Keys
        If Not dictLocal.Exists(k) Then
            If dictSnap.Exists(k) Then
                ' 取得時点にもあった → 管理者が削除した
                rowsToDelete.Add dictMaster(k)
                deleted = deleted + 1
            Else
                ' 取得時点には無かった → 他者が追加したので保持
                preserved = preserved + 1
            End If
        End If
    Next k

    ' 削除は後ろから行う（行番号がズレないように）
    Dim toDelete() As Long, n As Long: n = rowsToDelete.Count
    If n > 0 Then
        ReDim toDelete(1 To n)
        Dim j As Long
        For j = 1 To n
            toDelete(j) = rowsToDelete(j)
        Next j
        Call SortDescending(toDelete)
        For j = 1 To n
            masterSh.Rows(toDelete(j)).Delete
        Next j
    End If

    Application.CutCopyMode = False
    SyncSheetMerge = "編集:" & updated & " / 追加:" & added & _
                     " / 削除:" & deleted & " / 他者追加保持:" & preserved
End Function


'================================================================================
' ヘルパー: キー値 → 行番号 の辞書を作る
'================================================================================
Public Function BuildKeyMap(ByVal ws As Worksheet, ByVal keyCol As String, ByVal dataRow As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then Set BuildKeyMap = d: Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastRow < dataRow Then Set BuildKeyMap = d: Exit Function

    Dim r As Long, v As String
    For r = dataRow To lastRow
        v = Trim(CStr(ws.Cells(r, keyCol).Value))
        If v <> "" Then
            If Not d.Exists(v) Then d.Add v, r
        End If
    Next r
    Set BuildKeyMap = d
End Function


'================================================================================
' ヘルパー: 行コピー（値と書式）
'================================================================================
Public Sub CopyRow(ByVal src As Worksheet, ByVal srcRow As Long, ByVal dst As Worksheet, ByVal dstRow As Long, ByVal lastCol As Long)
    If lastCol < 1 Then lastCol = 1
    src.Range(src.Cells(srcRow, 1), src.Cells(srcRow, lastCol)).Copy _
        Destination:=dst.Cells(dstRow, 1)
End Sub


'================================================================================
' ヘルパー: 対象シートのキー列で最終データ行を得る
'================================================================================
Public Function GetLastDataRow(ByVal ws As Worksheet, ByVal keyCol As String) As Long
    GetLastDataRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
End Function


'================================================================================
' ヘルパー: Long配列を降順ソート
'================================================================================
Public Sub SortDescending(arr() As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) > arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub


'================================================================================
' ヘルパー: srcシートの全内容をスナップショットシートに書く
'================================================================================
Public Sub WriteSnapshotFromSheet(ByVal src As Worksheet, ByVal baseName As String)
    Dim snapName As String: snapName = SNAPSHOT_PREFIX & baseName
    Dim snapSh As Worksheet

    ' 既存スナップシートを取得 or 作成
    If SheetExistsIn(ThisWorkbook, snapName) Then
        Set snapSh = ThisWorkbook.Sheets(snapName)
    Else
        Set snapSh = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        snapSh.Name = snapName
    End If

    On Error Resume Next
    snapSh.Visible = xlSheetVeryHidden  ' 非表示（管理者からも見えない）
    snapSh.Cells.Clear
    If src.UsedRange.Cells.CountLarge > 0 Then
        src.UsedRange.Copy snapSh.Range("A1")
    End If
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub


'================================================================================
' 反映直後に静かに最新取得（メッセージ無し、スナップショット再作成が目的）
'================================================================================
Public Sub SilentRefreshSnapshots()
    Dim masterPath As String: masterPath = GetMasterPath()
    If Dir(masterPath) = "" Then Exit Sub

    Dim wbMaster As Workbook
    On Error GoTo Cleanup
    Set wbMaster = Workbooks.Open(fileName:=masterPath, ReadOnly:=True, UpdateLinks:=0)

    Dim configs As Variant: configs = GetSheetSyncConfig()
    Dim i As Long
    For i = LBound(configs) To UBound(configs)
        Dim sheetName As String: sheetName = CStr(configs(i)(0))
        If SheetExistsIn(wbMaster, sheetName) Then
            Call WriteSnapshotFromSheet(wbMaster.Sheets(sheetName), sheetName)
        End If
    Next i

    wbMaster.Close SaveChanges:=False
    Exit Sub

Cleanup:
    If Not wbMaster Is Nothing Then wbMaster.Close SaveChanges:=False
End Sub
