Attribute VB_Name = "保存"
Option Explicit

'================================================================================
' [保存] 保存（「保存」ボタンの本体 → 請求書を別 .xlsx として書き出し）
'--------------------------------------------------------------------------------
' 【何のファイル？】
'   作成済み請求書シートを、別ファイル(.xlsx)として
'   「YYYY年M月請求書」フォルダ配下に書き出す処理。
'
' 【しくみ】
'   - シートを Copy After:= でテンプレ後ろに複製してから新ブックへ
'     （Copy Before:= だと前提が崩れて空ブックになるバグの修正済み）
'   - AJ列以降のボタン図形を削除（コピーしたシートに残らないように）
'   - 出力ファイル名にテンプレ名を付けて衝突回避
'================================================================================


'================================================================================
' モジュール: 保存（【共有_経理用】請求書作成.xlsm 用）
'================================================================================
' 【役割】
'   請求書シート（No_xx_請求書 等）の保存処理。
'   - シートを xlsx として独立ファイルに保存
'   - 同時にマスタの依頼履歴へ経理番号を書き戻し
'
' 【ボタン登録】
'   各テンプレート請求書シートの右側（AJ列以降）にある「保存」ボタンに
'   登録されている。保存実行時は、その請求書シートが ActiveSheet になる。
'
' 【保存パス構造】
'   依頼検索!C1（ベース）\YYYY年M月請求書\<AK1のファイル名>.xlsx
'
'   - ベースフォルダ: 依頼検索!C1 セルで指定（空ならDownloads）
'   - 年月フォルダ : ActiveSheet の AB3(年) + AE3(月) から構築
'   - ファイル名   : ActiveSheet の AK1 から取得
'
' 【依存関係】
'   - Config モジュール: GetMasterPath(), SHEET_IRAI_RIREKI, SHEET_IRAI_SEARCH
'
' 【重要なバグ修正履歴】
'   - 2026/5/20: シートコピー Before→After に修正（保存xlsxがまっさらになるバグ）
'   - 2026/5/20: AJ列以降のシェイプ削除（保存ボタンが xlsx に残るバグ）
'   - 2026/5/18: 月年フォルダ名を「2026年4月」順に修正（旧: 4月2026 だった）
'   - 2026/5/18: 依頼検索!C1 から保存パス取得に変更（旧: 環境変数固定だった）
'
' 【セル位置定数の解説】
'   AK1 = 保存ファイル名（請求書作成マクロが書き込む。例: "1_請求書_㈱○○_工事名_03-2600001"）
'   AF1 = 経理番号（マスタへ書き戻す対象）
'   AB3 = 年（請求書日付の年。月年フォルダ名に使う）
'   AE3 = 月（請求書日付の月。月年フォルダ名に使う）
'================================================================================


' --- セル位置定数（請求書シート上の重要セル） ---
Private Const CELL_INVOICE_FILE_NAME As String = "AK1"   ' 保存ファイル名
Private Const CELL_ACCOUNTING_NO     As String = "AF1"   ' 経理番号
Private Const CELL_MONTH_COL1        As String = "AE3"   ' 月
Private Const CELL_MONTH_COL2        As String = "AB3"   ' 年


'================================================================================
' 主要処理：保存＆印刷（ボタンから呼び出される）
'================================================================================
' 【処理フロー】
'   1. マスタ同期       : SyncInvoiceNoToMaster で経理No をマスタに書き込み
'   2. ファイル名取得   : ActiveSheet!AK1 から取得
'   3. 年月取得         : ActiveSheet!AB3, AE3 から「YYYY年M月」を構築
'   4. 保存先決定       : 依頼検索!C1（無ければ Downloads）+ 年月フォルダ
'   5. フォルダ作成     : FileSystemObject で確実に作成
'   6. 新ブック作成     : ActiveSheet を新ブックに After:= でコピー
'   7. 不要シート削除   : 元の空 Sheet1 を削除
'   8. ボタン削除       : AJ列以降のシェイプを削除（保存ボタン等を除く）
'   9. xlsx保存         : SaveAs → Close
'  10. 完了メッセージ   : 保存パスをユーザーに通知
'
' 【エラー時】
'   On Error GoTo Cleanup で診断情報付きメッセージを表示
'   ScreenUpdating は必ず元に戻す
'================================================================================
Public Sub 保存_印刷作業()
    Dim activeSht As Worksheet
    Set activeSht = ActiveSheet
    Dim initialScreen As Boolean
    initialScreen = Application.ScreenUpdating

    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    ' --- ① マスタ同期：経理番号をマスタの依頼履歴へ書き込み ---
    Call SyncInvoiceNoToMaster

    Dim baseFolder As String, invoiceFolder As String
    Dim fileName As String, monthYear As String

    ' --- ② ファイル名取得（AK1）---
    fileName = CStr(activeSht.Range(CELL_INVOICE_FILE_NAME).Value)

    ' --- ③ 年月フォルダ名を「YYYY年M月」形式で構築 ---
    '   旧: 月+年 のおかしな順序だった不具合を修正
    Dim yearVal As String:  yearVal  = CStr(activeSht.Range(CELL_MONTH_COL2).Value)  ' AB3 = 年
    Dim monthVal As String: monthVal = CStr(activeSht.Range(CELL_MONTH_COL1).Value)  ' AE3 = 月
    If yearVal = "" Then yearVal = CStr(Year(Date))    ' 空欄なら今日の年でフォールバック
    If monthVal = "" Then monthVal = CStr(Month(Date)) ' 空欄なら今日の月でフォールバック
    monthYear = yearVal & "年" & monthVal & "月"

    ' --- ④ ファイル名に使えない文字(\/:*?<>|)を除去 ---
    fileName = SanitizeFileName(fileName)
    monthYear = SanitizeFileName(monthYear)

    If Trim(fileName) = "" Then
        MsgBox "ファイル名が取得できません。" & vbCrLf & _
               "請求書シートのAK1セルにファイル名が設定されているか確認してください。", vbCritical
        Exit Sub
    End If

    ' --- ⑤ 保存先ベースフォルダ取得 ---
    '   【2026/6/5 変更】テストモード対応:
    '     IS_TEST_MODE = True  → デスクトップ「請求書テスト」フォルダ（開発・動作確認用。
    '                            本番の保存先C1が開発環境に無くてもテストできる）
    '     IS_TEST_MODE = False → 依頼検索!C1 の値（空なら Downloads）… 本番動作
    If IS_TEST_MODE Then
        baseFolder = GetTestSaveFolder()
    Else
        On Error Resume Next
        baseFolder = Trim(CStr(ThisWorkbook.Sheets("依頼検索").Range("C1").Value))
        On Error GoTo Cleanup

        ' 両端のダブルクォート除去（C1 に "..." で入力されている場合の対策）
        Do While Len(baseFolder) > 0 And (Left(baseFolder, 1) = Chr(34) Or Left(baseFolder, 1) = "“" Or Left(baseFolder, 1) = "”")
            baseFolder = Mid(baseFolder, 2)
        Loop
        Do While Len(baseFolder) > 0 And (Right(baseFolder, 1) = Chr(34) Or Right(baseFolder, 1) = "“" Or Right(baseFolder, 1) = "”")
            baseFolder = Left(baseFolder, Len(baseFolder) - 1)
        Loop
        baseFolder = Trim(baseFolder)

        ' C1 が空なら Downloads にフォールバック
        If baseFolder = "" Then baseFolder = Environ("USERPROFILE") & "\Downloads"
    End If

    ' 末尾が \ なら削除（連結時の二重 \ 防止）
    If Right(baseFolder, 1) = "\" Then baseFolder = Left(baseFolder, Len(baseFolder) - 1)

    invoiceFolder = baseFolder & "\" & monthYear & "請求書"

    ' --- ⑥ FileSystemObject でフォルダ確実作成 ---
    '   OneDrive等の Known Folder Move 環境では MkDir が失敗するため FSO 必須
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseFolder) Then
        On Error Resume Next
        fso.CreateFolder baseFolder
        On Error GoTo Cleanup
    End If
    If Not fso.FolderExists(invoiceFolder) Then
        On Error Resume Next
        fso.CreateFolder invoiceFolder
        On Error GoTo Cleanup
    End If
    If Not fso.FolderExists(invoiceFolder) Then
        MsgBox "保存先フォルダが作成できませんでした。" & vbCrLf & _
               "パス: " & invoiceFolder & vbCrLf & vbCrLf & _
               "Downloadsフォルダの状態を確認してください。", vbCritical
        Exit Sub
    End If

    ' --- ⑦ 新ブック作成＆シートコピー ---
    Dim newWb As Workbook
    Set newWb = Workbooks.Add
    ' 【バグ修正 2026/5/20】
    '   旧: Copy Before:=Sheets(1) → Sheets(1).Delete
    '     これだとコピーした請求書シートが Sheets(1) になってしまい、
    '     Sheets(1).Delete が「コピーした請求書」を削除する致命的バグだった。
    '   新: Copy After:=Sheets(1) で末尾にコピー → 元の空Sheet1(=Sheets(1))を削除。
    activeSht.Copy After:=newWb.Sheets(1)
    Application.DisplayAlerts = False  ' シート削除の確認ダイアログを抑制
    newWb.Sheets(1).Delete  ' 元のSheet1（空）を削除。コピーした請求書は残る。
    Application.DisplayAlerts = True

    ' --- ⑦-2 明細書・納品書の同梱（任意）【2026/6/4 追加・お客様要望】 ---
    '   月1回程度、請求書と同時に明細書・納品書を発行する運用があるため、
    '   保存時に「同じファイルへ含めるか」を確認する（既定は「いいえ」）。
    '   含める場合は、このブック内の明細書・納品書シートを
    '   【手入力された現在の内容のまま】保存ファイルへコピーする。
    '   シートは「明細書」「納品書」で始まる名前を自動検出
    '   （「明細書 (追加分)」のような名前にも対応）。
    Dim wsMeisai As Worksheet, wsNouhin As Worksheet
    Set wsMeisai = FindSheetByPrefix("明細書")
    Set wsNouhin = FindSheetByPrefix("納品書")
    If Not (wsMeisai Is Nothing And wsNouhin Is Nothing) Then
        If MsgBox("明細書・納品書も同じファイルに含めますか？" & vbCrLf & vbCrLf & _
                  "（はい＝請求書・明細書・納品書をまとめて保存" & vbCrLf & _
                  "　いいえ＝請求書のみ保存）", _
                  vbYesNo + vbQuestion + vbDefaultButton2, "明細書・納品書の同梱") = vbYes Then
            Application.DisplayAlerts = False
            If Not wsMeisai Is Nothing Then
                wsMeisai.Copy After:=newWb.Sheets(newWb.Sheets.count)
                On Error Resume Next
                newWb.Sheets(newWb.Sheets.count).Name = "明細書"
                On Error GoTo Cleanup
            End If
            If Not wsNouhin Is Nothing Then
                wsNouhin.Copy After:=newWb.Sheets(newWb.Sheets.count)
                On Error Resume Next
                newWb.Sheets(newWb.Sheets.count).Name = "納品書"
                On Error GoTo Cleanup
            End If
            Application.DisplayAlerts = True
        End If
    End If

    ' --- ⑧ AJ列以降のボタン・図形を削除 ---
    '   テンプレ請求書の右側（AJ列以降）には「保存」ボタン等のマクロ用
    '   オブジェクトが配置されている。これが xlsx に含まれてお客様に渡ると
    '   不要なので、保存前に削除する。
    Dim copiedSht As Worksheet
    Set copiedSht = newWb.Sheets(1)
    Dim ajLeft As Double: ajLeft = copiedSht.Range("AJ1").Left
    Dim shp As Object
    For Each shp In copiedSht.Shapes
        On Error Resume Next
        If shp.Left >= ajLeft Then shp.Delete
        On Error GoTo 0
    Next shp

    ' --- ⑨ xlsx 保存 ---
    Dim filePath As String
    filePath = invoiceFolder & "\" & fileName & ".xlsx"
    newWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    ' --- ⑩ 完了 ---
    '   ※ 保存後の請求書シート（No_xx_請求書）はこのブック内に残す
    '     （印刷し直しや再確認のため。不要になったらユーザー側で手動削除）

    Application.ScreenUpdating = initialScreen
    MsgBox "ファイルを保存しました。" & vbCrLf & "保存先：" & filePath, vbInformation
    Exit Sub

Cleanup:
    Application.ScreenUpdating = initialScreen
    Application.DisplayAlerts = True  ' 同梱処理中のエラーでFalseのまま残るのを防止【2026/6/4】
    Dim diag As String
    diag = "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "【エラー詳細】" & vbCrLf & _
           " 番号: " & Err.Number & vbCrLf & _
           " 内容: " & Err.Description & vbCrLf & vbCrLf & _
           "【処理中の値】" & vbCrLf & _
           " ファイル名(AK1): [" & fileName & "]" & vbCrLf & _
           " 月年(AE3&AB3): [" & monthYear & "]" & vbCrLf & _
           " 保存フォルダ: [" & invoiceFolder & "]" & vbCrLf & _
           " 保存パス: [" & filePath & "]"
    MsgBox diag, vbCritical
End Sub


'================================================================================
' SyncInvoiceNoToMaster: マスタ同期処理
'================================================================================
' 【役割】
'   現在の請求書シート（ActiveSheet）の AF1（経理番号）を、
'   マスタファイルの「依頼履歴」シートのB列に書き込む。
'   検索キーは「依頼検索!A2」の依頼NO。
'
' 【副作用】
'   - マスタファイルを開く（書き込み権限）→ 書き込み後に Save → Close
'   - エラー時も必ず Cleanup ラベルを通って wbMaster.Close を実行
'
' 【エラー条件】
'   - マスタファイルのパスが空 → メッセージ表示してスキップ
'   - マスタファイルが存在しない → 同上
'   - マスタにロックがかかっている → Open で例外、Cleanup へ
'   - 依頼NOがマスタに存在しない → メッセージ表示してスキップ
'================================================================================
Private Sub SyncInvoiceNoToMaster()
    Dim mPath As String
    Dim wsSearch As Worksheet
    Dim wbMaster As Workbook
    Dim iraiNo As String
    Dim wsRireki As Worksheet
    Dim found As Range
    Dim targetRow As Long
    Dim accountingNo As String

    Set wsSearch = ThisWorkbook.Sheets(SHEET_IRAI_SEARCH)

    ' Config モジュールから直接マスタパスを取得（旧: Application.Run でラップしていた）
    mPath = GetMasterPath()

    On Error GoTo Cleanup

    ' 空パス対策：Dir("") はエラー52(ファイル名が不正)を投げるため事前チェック
    If Trim(mPath) = "" Then
        MsgBox "マスタファイルのパスが設定されていません。" & vbCrLf & _
               "依頼履歴シートのG1セルにパスを設定してください。", vbCritical
        GoTo Cleanup
    End If
    Dim dirChk As String
    On Error Resume Next
    dirChk = Dir(mPath)
    On Error GoTo Cleanup
    If dirChk = "" Then
        MsgBox "マスタファイルが見つかりません。" & vbCrLf & "パス: " & mPath, vbCritical
        GoTo Cleanup
    End If

    Set wbMaster = Workbooks.Open(mPath)
    If wbMaster Is Nothing Then
        MsgBox "マスタファイルを開けません。", vbCritical
        GoTo Cleanup
    End If

    iraiNo = Trim(wsSearch.Range("A2").Value)
    If iraiNo = "" Then
        MsgBox "依頼NOが指定されていません。", vbCritical
        GoTo Cleanup
    End If

    Set wsRireki = wbMaster.Sheets(SHEET_IRAI_RIREKI)
    Set found = wsRireki.Columns("A").Find(What:=iraiNo, LookIn:=xlValues, LookAt:=xlWhole)
    If found Is Nothing Then
        MsgBox "マスタ内に依頼NO [" & iraiNo & "] が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    targetRow = found.Row
    accountingNo = ActiveSheet.Range(CELL_ACCOUNTING_NO).Value
    ' 【シート保護対応 2026/5/20】マスタの依頼履歴が保護されている場合に備え、解除→書込→再保護
    Call SafeUnprotect(wsRireki)
    wsRireki.Cells(targetRow, 2).Value = accountingNo   ' B列に経理No書き込み
    Call SafeProtect(wsRireki)

    wbMaster.Save

Cleanup:
    ' 正常時もエラー時も必ずここを通る（マスタファイルが残らないように）
    If Not wbMaster Is Nothing Then wbMaster.Close False
    If Err.Number <> 0 Then
        MsgBox "マスタ同期中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    End If
End Sub


'================================================================================
' SanitizeFileName: ファイル名に使えない文字を安全な記号で置換
'================================================================================
' Windowsのファイル名禁止文字 \/:*?"<>| を全て "_" に置換する。
' 工事名に「㈱」や「（）」が含まれる場合があるので、それらは禁止文字に該当しないので残る。
'================================================================================
Private Function SanitizeFileName(ByVal name As String) As String
    Dim bad As String, i As Long, ch As String
    bad = "\/:*?""<>|"
    Dim result As String: result = name
    For i = 1 To Len(bad)
        ch = Mid(bad, i, 1)
        result = Replace(result, ch, "_")
    Next i
    SanitizeFileName = Trim(result)
End Function

' シート名が指定の接頭辞で始まるシートを探す（半角/全角スペースは無視）
'   例: FindSheetByPrefix("明細書") は「明細書」「明細書 (追加分)」等にマッチ
Private Function FindSheetByPrefix(ByVal prefix As String) As Worksheet
    Dim ws As Worksheet
    Dim normName As String
    For Each ws In ThisWorkbook.Worksheets
        normName = Replace(Replace(ws.Name, " ", ""), "　", "")
        If Left(normName, Len(prefix)) = prefix Then
            Set FindSheetByPrefix = ws
            Exit Function
        End If
    Next ws
    Set FindSheetByPrefix = Nothing
End Function
