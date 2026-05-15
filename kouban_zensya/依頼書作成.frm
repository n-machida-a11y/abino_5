VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 依頼書作成 
   Caption         =   "依頼書作成フォーム"
   ClientHeight    =   13836
   ClientLeft      =   24
   ClientTop       =   72
   ClientWidth     =   11136
   OleObjectBlob   =   "依頼書作成.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "依頼書作成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' IS_TEST_MODE / TEST_FILE_PATH / SHEET_* / CELL_* / PATH_CELL は Config モジュールで一元管理。

'--- 他のフォームから値を受け取るための変数 ---
Public m_SearchedKoujiName As String ' 検索対象の「工事名称」
Public m_PassedTantousha As String  ' 検索対象の「担当者名」

'--- このフォーム内だけで使う変数 ---
' ドロップダウンリスト同士が無限に更新しあうのを防ぐためのスイッチ（フラグ）
Private m_isComboBoxUpdating As Boolean
' 更新対象のデータの「行番号」を保存しておく変数
Private m_TargetRow As Long
' フォームに表示しているデータの「工事番号」を保存しておく変数
Private m_KoujiBangou As String

'--- 高速化のため、各種リストを一時保存するキャッシュ変数 ---
Private m_CachedSeikyuusakiList As Variant      ' 「請求書提出先」リスト

'--- セル位置マッピング（依頼書セル設定シートから読込） ---
Private m_CellMap As Object

'--- このフォームで処理する外部ファイルのパスを保存する変数 ---
Private m_TARGET_FILE_PATH As String

'--- プログラム内で固定的に使う文字（定数） ---
' 外部ファイル内のシート名（Config モジュールの定数を使用）
' SHEET_KOUJI_LIST   = "工事番号一覧"  (旧 SHEET_KOUJI_LIST)
' SHEET_KANRI_MASTER = "管理マスタ"    (旧 SHEET_KANRI_MASTER)
' SHEET_OTHER_MASTER = "その他マスタ"  (旧 SHEET_OTHER_MASTER)
' SHEET_IRAI_RIREKI  = "依頼履歴"      (旧 SHEET_IRAI_RIREKI)
' 「その他マスタ」シート内の列定義
Private Const MASTER_OTHER_SEIKYUUSAKI_COL As String = "A"      ' 請求書提出先
Private Const MASTER_OTHER_YUBIN_NO_COL As String = "B"         ' 郵便番号
Private Const MASTER_OTHER_JUSHO_COL As String = "C"            ' 住所

'--- 明細5行の固定行数（VBA側のフォーム対応行数） ---
Private Const MEISAI_ROW_COUNT As Long = 5



'================================================================================
' ★★★ 他のモジュールからこのフォームを呼び出すための専用プロシージャ ★★★
' 外部から .Show で直接表示するのではなく、このプロシージャを経由して呼び出します。
'================================================================================
Public Sub SetupAndShow(ByVal KoujiName As String, ByVal Tantousha As String)
    ' --- ① 他のモジュールから工事名と担当者名を受け取る ---
    Me.m_SearchedKoujiName = KoujiName
    Me.m_PassedTantousha = Tantousha

    ' --- ② フォームが表示される前に、初期値を設定する ---
    Me.作成日.Value = Format(Date, "yyyy/mm/dd")
    Me.担当者.Value = Me.m_PassedTantousha
    Me.工事名称.Value = Me.m_SearchedKoujiName
    Me.提出日付.Value = Format(Date, "yyyy/mm/dd")
    Me.数量1.Value = 1
    Me.単位1.Value = "式"
    Me.小計.Value = 0
    Me.消費税.Value = 0
    Me.請求金額.Value = 0

    ' --- ③ 初期値設定後に、フォームを表示する ---
    Me.Show
End Sub

Private Sub Label13_Click()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label35_Click()

End Sub

Private Sub Label36_Click()

End Sub

Private Sub Label41_Click()

End Sub

Private Sub Label45_Click()

End Sub

Private Sub Label46_Click()

End Sub

Private Sub Label47_Click()

End Sub

Private Sub Label57_Click()

End Sub

Private Sub Label58_Click()

End Sub

Private Sub Label59_Click()

End Sub

Private Sub Label62_Click()

End Sub

Private Sub Label67_Click()

End Sub

Private Sub Label68_Click()

End Sub

Private Sub Label73_Click()

End Sub

Private Sub TextBox14_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

'================================================================================
' フォームが開かれる瞬間の準備処理（１回だけ実行）
' ここでは、ドロップダウンリストの中身など、フォームの基本的な部品を準備します。
'================================================================================

Private Sub UserForm_Terminate()
    ' メモリリーク防止：オブジェクト参照を明示的に解放
    Set m_CellMap = Nothing
    m_CachedSeikyuusakiList = Empty
End Sub

Private Sub UserForm_Initialize()
    Dim wbTarget_Init As Workbook
    Dim wsKanriMaster As Worksheet, wsOtherMaster As Worksheet
    Dim originalDisplayAlerts As Boolean, originalEnableEvents As Boolean

    ' --- 処理中の画面のちらつきや不要なメッセージを抑制 ---
    Application.ScreenUpdating = False
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    originalEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    m_TargetRow = 0 ' 更新対象の行番号をリセット
    Me.請求書提出先.MatchEntry = fmMatchEntryNone ' ドロップダウンの入力補完を無効化（自前のロジックを使うため）

    On Error GoTo ErrorHandlerInit

    ' --- キャッシュデータがあれば、ファイルを開かずに高速化 ---
    If Not IsEmpty(m_CachedSeikyuusakiList) Then GoTo FinalizeInitWithoutFileOpen

    ' --- 外部ファイルを開き、ドロップダウンリスト用のマスタデータを読み込む ---
    m_TARGET_FILE_PATH = GetTargetFilePath()

    If Dir(m_TARGET_FILE_PATH) = "" Then GoTo FinalizeInit

    Set wbTarget_Init = Application.Workbooks.Open(fileName:=m_TARGET_FILE_PATH, ReadOnly:=True, UpdateLinks:=0)

    If Not SheetExists(wbTarget_Init, SHEET_KANRI_MASTER) Or Not SheetExists(wbTarget_Init, SHEET_OTHER_MASTER) Then GoTo FinalizeInit
    Set wsKanriMaster = wbTarget_Init.Sheets(SHEET_KANRI_MASTER)
    Set wsOtherMaster = wbTarget_Init.Sheets(SHEET_OTHER_MASTER)

    ' --- 各種マスタデータを読み込み、キャッシュに保存 ---
    m_CachedSeikyuusakiList = GetColumnData(wsOtherMaster, MASTER_OTHER_SEIKYUUSAKI_COL, 2)
    Me.請求書提出先.List = m_CachedSeikyuusakiList


    ' 依頼書セル設定シートからセル位置マッピングを読み込み
    Call LoadCellSettings(wbTarget_Init)

    ' --- PDF作成用のNamed Rangeをフォーム初期化時点で登録する ---
    '     これにより「フォームを開くだけ→PDF実行」でもNamed Rangeが使える
    On Error Resume Next
    Dim wsReqEarly As Worksheet
    Set wsReqEarly = ThisWorkbook.Sheets(SHEET_IRAISHO)
    If Not wsReqEarly Is Nothing Then
        Call RegisterSheetCellName(ThisWorkbook, wsReqEarly, "PDF_請求宛名", GetCellAddr("請求宛名"))
        Call RegisterSheetCellName(ThisWorkbook, wsReqEarly, "PDF_工事名称", GetCellAddr("工事名称"))
    End If
    Err.Clear
    On Error GoTo 0

    ' 提出要項・同封物はOptionButtonに変更済み（マスタ読み込み不要）

FinalizeInit: ' 終了処理（ファイルを開いた場合）
    If Not wbTarget_Init Is Nothing Then wbTarget_Init.Close SaveChanges:=False

FinalizeInitWithoutFileOpen: ' 終了処理（ファイルを開かなかった場合も含む）
    ' 作成日を編集不可にする
    Me.作成日.Enabled = False
    
    ' 小計と消費税は自動計算なので、手入力できないようにする
    Me.小計.Enabled = False
    Me.消費税.Enabled = False

    ' --- Excelの設定を元に戻す ---
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandlerInit: ' エラー発生時の処理
    Resume FinalizeInit
End Sub

'================================================================================
' フォームがアクティブになった時の処理（特定の工事データを検索し、フォームに表示）
'================================================================================
Private Sub UserForm_Activate()
    Dim wbTarget_Activate As Workbook
    Dim wsTarget_Activate As Worksheet
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean

    ' --- 処理中の画面のちらつきや不要なメッセージを抑制 ---
    Application.ScreenUpdating = False
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    originalEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerActivate

    ' --- 編集対象のデータが指定されているか確認 ---
    If Trim(Me.m_SearchedKoujiName) = "" Or Trim(Me.m_PassedTantousha) = "" Then
        Unload Me
        GoTo FinalizeActivate
    End If

    ' Initializeで準備したキャッシュデータがあるか確認
    If IsEmpty(m_CachedSeikyuusakiList) Then
        Unload Me
        GoTo FinalizeActivate
    End If

    ' --- 外部ファイルを読み取り専用で開く ---
    Dim targetFilePath As String
    targetFilePath = GetTargetFilePath()

    If Dir(targetFilePath) = "" Then
        Unload Me
        GoTo FinalizeActivate
    End If

    Set wbTarget_Activate = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)

    If Not SheetExists(wbTarget_Activate, SHEET_KOUJI_LIST) Then
        GoTo FinalizeActivate
    End If
    Set wsTarget_Activate = wbTarget_Activate.Sheets(SHEET_KOUJI_LIST)

    ' --- データシートを１行ずつ調べ、工事名称と担当者の両方が一致する行を探す ---
    m_TargetRow = 0
    Dim r As Long
    Dim sheetKoujiName As String, sheetStaffName As String
    For r = 2 To wsTarget_Activate.Cells(wsTarget_Activate.Rows.count, "E").End(xlUp).Row
        sheetKoujiName = Trim(CStr(wsTarget_Activate.Cells(r, "E").Value))
        sheetStaffName = Trim(CStr(wsTarget_Activate.Cells(r, "C").Value))

        If sheetKoujiName = Trim(Me.m_SearchedKoujiName) And sheetStaffName = Trim(Me.m_PassedTantousha) Then
            m_TargetRow = r ' 一致する行が見つかった
            Exit For      ' ループを抜ける
        End If
    Next r

    ' --- 検索結果の処理 ---
    If m_TargetRow > 0 Then
        ' データが見つかった場合、工事番号を取得し、既存のデータをフォームに読み込む
        m_KoujiBangou = wsTarget_Activate.Cells(m_TargetRow, "D").Value
        Call LoadDataToForm(wsTarget_Activate, m_TargetRow)
    Else
        ' データが見つからなかった場合はフォームを閉じる
        Unload Me
        GoTo FinalizeActivate
    End If

    ' 金額の合計を再計算
    m_isComboBoxUpdating = False
    Call CalculateTotals

FinalizeActivate: ' 終了処理
    On Error Resume Next ' Unload Me後のアクセスに備える
    Me.小計.Enabled = False
    Me.消費税.Enabled = False
    On Error GoTo 0

    If Not wbTarget_Activate Is Nothing Then wbTarget_Activate.Close SaveChanges:=False
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandlerActivate: ' エラー発生時の処理
    Resume FinalizeActivate
End Sub

'================================================================================
' 「依頼書作成」ボタンがクリックされたときの処理
'================================================================================
Private Sub 依頼書作成_Click()
    Dim wbTarget_Click As Workbook
    Dim wsRequest As Worksheet, wsSrc As Worksheet, wsMaster_Click As Worksheet
    Dim wsRireki As Worksheet
    Dim wsRirekiLocal As Worksheet
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim isSuccess As Boolean

    isSuccess = False ' 処理成功フラグを初期化

    ' --- 処理中の画面のちらつきや不要なメッセージを抑制 ---
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerClick

    ' --- 事前チェック ---
    If m_TargetRow = 0 Then GoTo CleanUpClick ' 更新対象行が不明なら中断
    If MsgBox("依頼書シートを作成します。よろしいですか？", vbYesNo + vbQuestion, "確認") = vbNo Then GoTo CleanUpClick ' ユーザーがキャンセルしたら中断

    ' --- ① 外部ファイルを書き込みモードで開く ---
    Dim targetFilePath As String
    targetFilePath = GetTargetFilePath()

    ' ファイルが既に開かれていないかチェック（二重編集を防ぐため）
    Dim openedWbMain As Workbook
    For Each openedWbMain In Application.Workbooks
        If openedWbMain.FullName = targetFilePath Then
            MsgBox "対象のExcelファイルが既に開かれています。閉じてください。", vbCritical
            GoTo CleanUpClick
        End If
    Next openedWbMain

    On Error Resume Next
    Set wbTarget_Click = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=False, UpdateLinks:=0)
    On Error GoTo ErrorHandlerClick

    If wbTarget_Click Is Nothing Then
        MsgBox "対象のExcelファイルを開けませんでした。処理を中断します。", vbCritical
        GoTo CleanUpClick
    End If
    If wbTarget_Click.ReadOnly Then
        MsgBox "対象のExcelファイルは読み取り専用で開かれました。" & vbCrLf & _
               "他のユーザーが使用中の可能性があります。処理を中断します。", vbExclamation
        GoTo CleanUpClick
    End If

    ' 必要なシートが存在するかチェック
    If Not SheetExists(wbTarget_Click, SHEET_KOUJI_LIST) Or _
       Not SheetExists(wbTarget_Click, SHEET_OTHER_MASTER) Or _
       Not SheetExists(wbTarget_Click, SHEET_IRAI_RIREKI) Then
        MsgBox "必要なシート（工事番号一覧、その他マスタ、依頼履歴）のいずれかが見つかりません。", vbCritical
        GoTo CleanUpClick
    End If
    Set wsSrc = wbTarget_Click.Sheets(SHEET_KOUJI_LIST)
    Set wsMaster_Click = wbTarget_Click.Sheets(SHEET_OTHER_MASTER)
    Set wsRireki = wbTarget_Click.Sheets(SHEET_IRAI_RIREKI)

    ' --- ② フォームの入力内容で、外部ファイルのデータを更新 ---
    ' (2-1) 「その他マスタ」に、新しい請求先情報があれば追加・更新する
    Call UpdateAddressMaster(wsMaster_Click)
    
    ' (2-2) 「工事番号一覧」シートのN列以降を更新する
    Call UpdateExternalFile(wsSrc, m_TargetRow)

    ' (2-2) 「依頼履歴」シートに、今回の依頼内容を新しい行として追記する
    Call AddDataToIraiRireki(wsRireki)
    
    ' (2-4) 「工事番号一覧」の色付け処理を停止（ユーザーの要望により）
    ' With wsSrc.Rows(m_TargetRow).Interior
    '     .Pattern = xlSolid
    '     .Color = RGB(220, 220, 220)  ' 色を薄めのグレーに設定
    ' End With

    ' 更新内容を保存（共有ファイルのフィルタは保存前にクリア）
    Call ClearAllFilters(wsSrc)
    Call ClearAllFilters(wsRireki)
    Application.EnableEvents = False ' イベントを一時的に停止
    wbTarget_Click.Save              ' 保存を実行
    Application.EnableEvents = True  ' イベントを元に戻す

    ' --- ③ このツール内の「請求書提出依頼書」シートに内容を転記 ---
    Set wsRequest = ThisWorkbook.Sheets("請求書提出依頼書")
    Call SafeUnprotect(wsRequest)

    With wsRequest
        ' --- ヘッダー部 ---
        .Range(GetCellAddr("請求宛名")).Value = Me.請求書提出先.Value                    ' 請求宛名
        .Range(GetCellAddr("郵送先住所")).Value = "〒" & Me.郵便番号.Value & "　" & Me.住所.Value  ' 郵送先住所
        .Range("F7").Value = Val(Replace(Me.請求金額.Value, ",", ""))  ' 請求金額(税込)
        .Range(GetCellAddr("消費税テキスト")).Value = "（ 内消費税　　￥" & Me.消費税.Value & " ）"  ' 消費税テキスト
        .Range(GetCellAddr("工事名称")).Value = Me.工事名称.Value                        ' 工事名称
        .Range(GetCellAddr("工事番号")).Value = m_KoujiBangou                            ' 工事番号
        .Range(GetCellAddr("提出要項")).Value = GetSelectedTeishutsuyoukou()             ' 提出要項
        .Range(GetCellAddr("同封物")).Value = GetSelectedDoufuubutsu()                 ' 同封物
        
        ' --- 日付部（行12） ---
        If IsDate(Me.着手.Value) Then .Range(GetCellAddr("着手")).Value = CDate(Me.着手.Value)
        If IsDate(Me.完成.Value) Then .Range(GetCellAddr("完成")).Value = CDate(Me.完成.Value)
        If IsDate(Me.引渡日.Value) Then .Range(GetCellAddr("引渡日")).Value = CDate(Me.引渡日.Value)
        If IsDate(Me.提出日付.Value) Then .Range(GetCellAddr("請求書日付")).Value = CDate(Me.提出日付.Value)
        
        ' --- 明細5行（行14-18）---
        .Range("F14").Value = Me.txt名称1.Value
        .Range("N14").Value = Val(Me.数量1.Value)
        .Range("P14").Value = Me.単位1.Value
        .Range("Q14").Value = Val(Replace(Me.金額1.Value, ",", ""))
        .Range("R14").Value = Val(Replace(Me.金額1.Value, ",", ""))
        
        .Range("F15").Value = Me.txt名称2.Value
        .Range("N15").Value = Val(Me.数量2.Value)
        .Range("P15").Value = Me.単位2.Value
        .Range("Q15").Value = Val(Replace(Me.金額2.Value, ",", ""))
        .Range("R15").Value = Val(Replace(Me.金額2.Value, ",", ""))
        
        .Range("F16").Value = Me.txt名称3.Value
        .Range("N16").Value = Val(Me.数量3.Value)
        .Range("P16").Value = Me.単位3.Value
        .Range("Q16").Value = Val(Replace(Me.金額3.Value, ",", ""))
        .Range("R16").Value = Val(Replace(Me.金額3.Value, ",", ""))
        
        .Range("F17").Value = Me.txt名称4.Value
        .Range("N17").Value = Val(Me.数量4.Value)
        .Range("P17").Value = Me.単位4.Value
        .Range("Q17").Value = Val(Replace(Me.金額4.Value, ",", ""))
        .Range("R17").Value = Val(Replace(Me.金額4.Value, ",", ""))
        
        .Range("F18").Value = Me.txt名称5.Value
        .Range("N18").Value = Val(Me.数量5.Value)
        .Range("P18").Value = Me.単位5.Value
        .Range("Q18").Value = Val(Replace(Me.金額5.Value, ",", ""))
        .Range("R18").Value = Val(Replace(Me.金額5.Value, ",", ""))
        
        ' --- 小計・消費税（行19-20）---
        .Range("R19").Value = Val(Replace(Me.小計.Value, ",", ""))
        .Range("R20").Value = Val(Replace(Me.消費税.Value, ",", ""))
        
        ' --- 引継ぎコメント ---
        .Range(GetCellAddr("引継ぎコメント")).Value = Me.引継ぎコメント.Value
        
        ' --- 作成日（年・月・日を分割）---
        If IsDate(Me.作成日.Value) Then
            .Range(GetCellAddr("作成日(年)")).Value = Year(CDate(Me.作成日.Value))
            .Range(GetCellAddr("作成日(月)")).Value = Month(CDate(Me.作成日.Value))
            .Range(GetCellAddr("作成日(日)")).Value = Day(CDate(Me.作成日.Value))
        End If
        
        ' --- 担当者名 ---
        .Range(GetCellAddr("担当者名")).Value = Me.担当者.Value

        ' --- 提出依頼者所属部署（各管理Excelごとに設定用コード.basで固定値を設定） ---
        If SUBMITTER_DEPARTMENT <> "" Then
            Dim depCell As String: depCell = GetCellAddr("提出依頼者所属部署")
            If depCell <> "" Then .Range(depCell).Value = SUBMITTER_DEPARTMENT
        End If

        ' --- PDF作成用: ファイル名に使うセル位置を Named Range で記憶 ---
        '     (PDF作成.bas がレイアウト変更に追従できるように)
        Call RegisterSheetCellName(ThisWorkbook, wsRequest, "PDF_請求宛名", GetCellAddr("請求宛名"))
        Call RegisterSheetCellName(ThisWorkbook, wsRequest, "PDF_工事名称", GetCellAddr("工事名称"))
    End With
    Call SafeProtectFull(wsRequest)

    ' --- ④ このツール内の工事一覧シートも最新の状態に更新 ---
    Call UpdateLocalListSheet(wsSrc, wbTarget_Click.Sheets(SHEET_KANRI_MASTER))

    ' --- ⑤ このツール内の「依頼履歴」シートも最新の状態に更新 ---
    On Error Resume Next
    Set wsRirekiLocal = ThisWorkbook.Sheets(SHEET_IRAI_RIREKI)
    On Error GoTo ErrorHandlerClick
    
    If wsRirekiLocal Is Nothing Then
        MsgBox "このファイルに「" & SHEET_IRAI_RIREKI & "」が見つかりませんでした。", vbExclamation
    Else
        Call UpdateLocalRirekiSheet(wsRireki, wsRirekiLocal)
    End If

    isSuccess = True ' 全ての処理が成功

CleanUpClick: ' 終了処理
    If Not wbTarget_Click Is Nothing Then wbTarget_Click.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    If isSuccess Then MsgBox "依頼書と一覧シートを更新しました。", vbInformation, "完了"
    Unload Me
    Exit Sub

ErrorHandlerClick: ' エラー発生時の処理
    Resume CleanUpClick
End Sub

Private Sub 契約無し_Click()

End Sub

'================================================================================
' フォーム上のコントロールに対するイベント処理
'================================================================================

' 「税込み」ボタンが押されたときの処理
Private Sub 税込み_Click()
    Dim val1 As Currency, val2 As Currency, val3 As Currency, val4 As Currency, val5 As Currency
    Dim subTotal As Currency
    
    ' 金額1～5の合計を算出（空欄は0扱い）
    val1 = Val(Replace(Me.金額1.Value, ",", ""))
    val2 = Val(Replace(Me.金額2.Value, ",", ""))
    val3 = Val(Replace(Me.金額3.Value, ",", ""))
    val4 = Val(Replace(Me.金額4.Value, ",", ""))
    val5 = Val(Replace(Me.金額5.Value, ",", ""))
    subTotal = val1 + val2 + val3 + val4 + val5
    
    If subTotal = 0 Then
        MsgBox "金額が入力されていません。", vbExclamation
        Exit Sub
    End If
    
    ' 税込み合計として小計＝請求金額とし、消費税は0にする（内訳はそのまま保持）
    Me.小計.Value = Format(subTotal, "#,##0")
    Me.消費税.Value = 0
    Me.請求金額.Value = Format(subTotal, "#,##0")
End Sub

' 「請求書提出先」ドロップダウンの入力が変更されたときの処理（オートコンプリート機能）
Private Sub 請求書提出先_Change()
    If m_isComboBoxUpdating Then Exit Sub ' 無限ループ防止
    Dim originalText As String, selStart As Long
    Dim filteredList As Object, item As Variant
    Dim wbTarget_AutoFill As Workbook, wsMaster_AutoFill As Worksheet
    On Error GoTo ErrorHandlerAutoFill

    ' --- ① 入力された文字を含む候補だけにリストを絞り込む ---
    originalText = Me.請求書提出先.Text
    selStart = Me.請求書提出先.selStart
    Set filteredList = CreateObject("Scripting.Dictionary")
    If originalText <> "" And Not IsEmpty(m_CachedSeikyuusakiList) Then
        For Each item In m_CachedSeikyuusakiList
            If InStr(1, item, originalText, vbTextCompare) > 0 Then ' 大文字小文字を区別せず検索
                If Not filteredList.Exists(item) Then filteredList.Add item, True
            End If
        Next item
    End If

    ' --- ② 絞り込んだリストをドロップダウンに再設定 ---
    m_isComboBoxUpdating = True ' これからリストを更新することを示すスイッチON
    Me.請求書提出先.Clear
    If filteredList.count > 0 Then Me.請求書提出先.List = filteredList.Keys
    Me.請求書提出先.Text = originalText ' 入力途中の文字を復元
    Me.請求書提出先.selStart = selStart ' カーソル位置を復元
    If filteredList.count > 0 Then Me.請求書提出先.DropDown ' 候補リストを表示
    m_isComboBoxUpdating = False ' スイッチOFF

    ' --- ③ リストから項目が完全に選択されたら、関連情報を自動入力 ---
    If Me.請求書提出先.ListIndex > -1 Then
        Dim targetFilePath As String
        targetFilePath = GetTargetFilePath()
        Set wbTarget_AutoFill = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)
        Set wsMaster_AutoFill = wbTarget_AutoFill.Sheets(SHEET_OTHER_MASTER)
        ' マスタから郵便番号や住所などを探して自動入力する
        Call AutoFillFromMaster(Me.請求書提出先.Value, wsMaster_AutoFill)
        If Not wbTarget_AutoFill Is Nothing Then wbTarget_AutoFill.Close SaveChanges:=False
    End If
Exit Sub
ErrorHandlerAutoFill:
    MsgBox "請求書提出先自動入力中にエラー発生: " & Err.Description, vbCritical
End Sub

' 金額欄が変更されたら、合計を自動計算する
Private Sub 金額1_Change(): Call CalculateTotals: End Sub
Private Sub 金額2_Change(): Call CalculateTotals: End Sub
Private Sub 金額3_Change(): Call CalculateTotals: End Sub
Private Sub 金額4_Change(): Call CalculateTotals: End Sub
Private Sub 金額5_Change(): Call CalculateTotals: End Sub

' 日付欄からフォーカスが外れたら、書式をチェック・統一する
Private Sub 作成日_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.作成日, "作成日", Cancel): End Sub
Private Sub 着手_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.着手, "工期着手", Cancel): End Sub
Private Sub 完成_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.完成, "工期完成", Cancel): End Sub
Private Sub 引渡日_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.引渡日, "引渡日", Cancel): End Sub
Private Sub 提出日付_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.提出日付, "提出日付", Cancel): End Sub

'================================================================================
' 補助的な関数やサブルーチン群 (Helper Functions)
'================================================================================

' フォームの入力内容を、外部ファイルの「依頼履歴」シートに追記する
Private Sub AddDataToIraiRireki(ByVal wsRireki As Worksheet)
    If wsRireki Is Nothing Then Exit Sub

    On Error GoTo ErrorHandlerAddRireki
    Call SafeUnprotect(wsRireki)
    
    Dim nextRow As Long, lastRow As Long
    Dim lastNoStr As String, nextNoNum As Long, newIraiNo As String
    
    ' --- 1. 追記する行番号を決定 (1行目はヘッダー) ---
    lastRow = wsRireki.Cells(wsRireki.Rows.count, "A").End(xlUp).Row
    If lastRow < 1 Then ' シートが完全に空の場合
        nextRow = 2 ' 2行目から書き込む
    ElseIf lastRow = 1 Then ' ヘッダーしかない場合
        nextRow = 2 ' 2行目から書き込む
    Else
        nextRow = lastRow + 1
    End If
    
    ' --- 2. 依頼NOを自動採番 ("0301", "0302"...) ---
    ' 全行を走査して最大の連番を取得する。
    ' Excelが先頭ゼロを自動削除して数値保存する場合（例: "0301"→301）も考慮し
    ' 301?399の数値も旧形式（"03XX"の先頭ゼロ落ち）として検出する。
    Dim i As Long, maxNum As Long, cellVal As String, cellNum As Long, numVal As Long
    maxNum = 0
    For i = 2 To lastRow
        cellVal = Trim(CStr(wsRireki.Cells(i, "A").Value))
        If Left(cellVal, 2) = "03" And IsNumeric(Mid(cellVal, 3)) Then
            ' テキスト形式: "0301", "0302" など
            cellNum = CLng(Mid(cellVal, 3))
            If cellNum > maxNum Then maxNum = cellNum
        ElseIf IsNumeric(cellVal) Then
            ' 数値として保存された旧形式: 301?399（="0301"?"0399"の先頭ゼロ落ち）
            numVal = CLng(cellVal)
            If numVal >= 301 And numVal <= 399 Then
                cellNum = numVal - 300
                If cellNum > maxNum Then maxNum = cellNum
            End If
        End If
    Next i
    nextNoNum = maxNum + 1
    newIraiNo = "03" & Format(nextNoNum, "00")

    ' --- 3. データを書き込む ---
    With wsRireki
        .Cells(nextRow, "A").Value = newIraiNo ' 依頼NO (A列)
        .Cells(nextRow, "B").Value = "" ' 請求書番号（経理用） (B列)
        .Cells(nextRow, "C").Value = "" ' 請求書発行日 (C列)
        .Cells(nextRow, "D").Value = Me.担当者.Value ' 1:担当者 (D列)
        .Cells(nextRow, "E").Value = m_KoujiBangou ' 2:工事番号 (E列)
        .Cells(nextRow, "F").Value = Me.工事名称.Value ' 3:工事名称 (F列)
        .Cells(nextRow, "G").Value = FormatIfDate(Me.着手.Value) ' 4:工期 着手 (G列)
        .Cells(nextRow, "H").Value = FormatIfDate(Me.完成.Value) ' 5:工期 完成 (H列)
        .Cells(nextRow, "I").Value = Val(Replace(Me.請求金額.Value, ",", "")) ' 6:請負金額 (I列)
        .Cells(nextRow, "J").Value = FormatIfDate(Me.作成日.Value) ' 7:依頼書作成日 (J列)
        .Cells(nextRow, "K").Value = GetSelectedTeishutsuyoukou() ' 8:提出要項 (K列)
        
        .Cells(nextRow, "L").Value = "" ' 9:提出要項(その他) (L列) ※廃止
        
        .Cells(nextRow, "M").Value = GetSelectedDoufuubutsu() ' 10:同封物 (M列)
        .Cells(nextRow, "N").Value = FormatIfDate(Me.提出日付.Value) ' 11:提出日付 (N列)
        .Cells(nextRow, "O").Value = FormatIfDate(Me.引渡日.Value) ' 12:引渡日 (O列)
        .Cells(nextRow, "P").Value = Me.請求書提出先.Value ' 13:請求書提出先 (P列)
        .Cells(nextRow, "Q").Value = Me.郵便番号.Value ' 14:郵便番号 (Q列)
        .Cells(nextRow, "R").Value = Me.住所.Value ' 15:提出先住所 (R列)
        
        .Cells(nextRow, "S").Value = "" ' 16:請求書記載文言 (S列) ※廃止
        
        .Cells(nextRow, "T").Value = Me.引継ぎコメント.Value ' 17:担当者引継ぎコメント (T列)
        
            .Cells(nextRow, "U").Value = "" ' 18:領収書注意文 (U列) ※廃止
            .Cells(nextRow, "V").Value = "" ' 19:振込手数料注意文 (V列) ※廃止
        If Me.但陽信金口座指定.Value = True Then ' 20:但陽信金口座指定 (W列)
            .Cells(nextRow, "W").Value = "有"
        Else
            .Cells(nextRow, "W").Value = ""
        End If

        ' 明細（5行分）をJSON形式でZ列に保存
        .Cells(nextRow, "Z").Value = BuildMeisaiJSON()
    End With

    ' PDF作成時にR3へ依頼NOを転記できるよう、請求書提出依頼書シートに書き込む
    Dim wsReq As Worksheet
    On Error Resume Next
    Set wsReq = ThisWorkbook.Sheets("請求書提出依頼書")
    On Error GoTo 0
    Call SafeProtectData(wsRireki)
    
    If Not wsReq Is Nothing Then
        Call SafeUnprotect(wsReq)
        wsReq.Range(GetCellAddr("依頼NO")).Value = newIraiNo
        Call SafeProtectFull(wsReq)
    End If
    Exit Sub

ErrorHandlerAddRireki:
    ' エラー発生時も必ず保護を復元
    Call SafeProtectData(wsRireki)
    If Not wsReq Is Nothing Then Call SafeProtectFull(wsReq)
    MsgBox "依頼履歴の更新中にエラー: " & Err.Description, vbCritical
End Sub


' 日付の形式が正しいかチェックし、自動で書式を整える
Private Sub ValidateDate(ByVal DateField As MSForms.Control, ByVal FieldName As String, ByRef Cancel As MSForms.ReturnBoolean)
    Dim inputText As String
    inputText = Trim(DateField.Value)
    If inputText = "" Then Exit Sub
    If IsDate(inputText) Then
        DateField.Value = Format(CDate(inputText), "yyyy/mm/dd")
    Else
        MsgBox FieldName & " は「YYYY/MM/DD」形式で入力してください。", vbExclamation, "入力エラー"
        Cancel = True
    End If
End Sub

' 金額1～3を合計し、小計・消費税・請求金額を計算する
Private Sub CalculateTotals()
    Dim val1 As Currency, val2 As Currency, val3 As Currency, val4 As Currency, val5 As Currency
    Dim subTotal As Currency, tax As Currency, grandTotal As Currency
    val1 = Val(Replace(Me.金額1.Value, ",", ""))
    val2 = Val(Replace(Me.金額2.Value, ",", ""))
    val3 = Val(Replace(Me.金額3.Value, ",", ""))
    val4 = Val(Replace(Me.金額4.Value, ",", ""))
    val5 = Val(Replace(Me.金額5.Value, ",", ""))
    subTotal = val1 + val2 + val3 + val4 + val5
    tax = Application.WorksheetFunction.Round(subTotal * 0.1, 0) ' 消費税10%を計算（四捨五入）
    grandTotal = subTotal + tax
    Me.小計.Value = Format(subTotal, "#,##0")
    Me.消費税.Value = Format(tax, "#,##0")
    Me.請求金額.Value = Format(grandTotal, "#,##0")
End Sub

' フォームの入力内容を、外部ファイルの「工事番号一覧」シートに書き込む
Private Sub UpdateExternalFile(ByVal wsTarget As Worksheet, ByVal rowToUpdate As Long)
    If wsTarget Is Nothing Or rowToUpdate = 0 Then Exit Sub

    On Error GoTo ErrorHandlerExt
    Call SafeUnprotect(wsTarget)

    With wsTarget
        .Cells(rowToUpdate, "S").Value = Me.請求書提出先.Value
        .Cells(rowToUpdate, "G").Value = FormatIfDate(Me.着手.Value)
        .Cells(rowToUpdate, "H").Value = FormatIfDate(Me.完成.Value)
        .Cells(rowToUpdate, "K").Value = Val(Replace(Me.小計.Value, ",", ""))
        .Cells(rowToUpdate, "C").Value = Me.担当者.Value
        .Cells(rowToUpdate, "Q").Value = FormatIfDate(Me.提出日付.Value)
        .Cells(rowToUpdate, "O").Value = GetSelectedTeishutsuyoukou()
        
        .Cells(rowToUpdate, "P").Value = GetSelectedDoufuubutsu()
        
        .Cells(rowToUpdate, "R").Value = FormatIfDate(Me.引渡日.Value)
        .Cells(rowToUpdate, "T").Value = Me.郵便番号.Value
        .Cells(rowToUpdate, "U").Value = Me.住所.Value
        .Cells(rowToUpdate, "X").Value = Me.引継ぎコメント.Value
        .Cells(rowToUpdate, "N").Value = Date ' 最終更新日
    End With
    Call SafeProtectData(wsTarget)
    Exit Sub

ErrorHandlerExt:
    Call SafeProtectData(wsTarget)
    MsgBox "工事番号一覧の更新中にエラー: " & Err.Description, vbCritical
End Sub

' フォームの入力内容で、外部ファイルの「その他マスタ」を更新する
Private Sub UpdateAddressMaster(ByVal wsMaster As Worksheet)
    If wsMaster Is Nothing Then Exit Sub

    On Error GoTo ErrorHandlerAddrMaster
    Call SafeUnprotect(wsMaster)

    Dim searchVal As String, foundCell As Range
    searchVal = Trim(Me.請求書提出先.Value)
    If searchVal <> "" Then
        ' マスタに同じ請求先があるか探す
        Set foundCell = wsMaster.Columns(MASTER_OTHER_SEIKYUUSAKI_COL).Find(What:=searchVal, LookIn:=xlValues, LookAt:=xlWhole)
        If foundCell Is Nothing Then
            ' 見つからない場合：新しいデータとして最終行に追加
            Dim lastRow As Long
            lastRow = wsMaster.Cells(wsMaster.Rows.count, MASTER_OTHER_SEIKYUUSAKI_COL).End(xlUp).Row + 1
            wsMaster.Cells(lastRow, MASTER_OTHER_SEIKYUUSAKI_COL).Value = searchVal
            wsMaster.Cells(lastRow, MASTER_OTHER_YUBIN_NO_COL).Value = Me.郵便番号.Value
            wsMaster.Cells(lastRow, MASTER_OTHER_JUSHO_COL).Value = Me.住所.Value
        Else
            ' 見つかった場合：既存のデータを最新の情報で上書き
            wsMaster.Cells(foundCell.Row, MASTER_OTHER_YUBIN_NO_COL).Value = Me.郵便番号.Value
            wsMaster.Cells(foundCell.Row, MASTER_OTHER_JUSHO_COL).Value = Me.住所.Value
        
        End If
    End If

    ' 「その他マスタ」は重要なデータなので、最後に保護をかける
    Call SafeProtectFull(wsMaster)
    Exit Sub

ErrorHandlerAddrMaster:
    Call SafeProtectFull(wsMaster)
    MsgBox "その他マスタの更新中にエラー: " & Err.Description, vbCritical
End Sub

' 「その他マスタ」から選択された請求先に紐づく情報を探し、フォームに自動入力する
Private Sub AutoFillFromMaster(ByVal selectedValue As String, ByVal wsMaster As Worksheet)
    If selectedValue = "" Or wsMaster Is Nothing Then Exit Sub
    Dim foundCell As Range
    On Error GoTo CleanUpAutoFill
    Set foundCell = wsMaster.Columns(MASTER_OTHER_SEIKYUUSAKI_COL).Find(What:=selectedValue, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        m_isComboBoxUpdating = True ' 無限ループ防止
        Me.郵便番号.Value = wsMaster.Cells(foundCell.Row, MASTER_OTHER_YUBIN_NO_COL).Value
        Me.住所.Value = wsMaster.Cells(foundCell.Row, MASTER_OTHER_JUSHO_COL).Value
        m_isComboBoxUpdating = False
    End If
CleanUpAutoFill:
    Exit Sub
End Sub




' シートの指定された列のデータを配列として取得する
Private Function GetColumnData(ByVal ws As Worksheet, ByVal col As String, ByVal startRow As Long) As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    If lastRow >= startRow Then
        GetColumnData = ws.Range(col & startRow & ":" & col & lastRow).Value
    Else
        GetColumnData = Array() ' データがない場合は空の配列を返す
    End If
End Function

' ドロップダウンリストに値を設定する（リストにない場合は追加してから設定）
Private Sub SetComboBoxValue(ByVal cmb As MSForms.ComboBox, ByVal valueToSet As String)
    Dim i As Long, found As Boolean
    If Trim(valueToSet) = "" Then Exit Sub
    found = False
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = valueToSet Then
            found = True
            Exit For
        End If
    Next i
    If found Then
        cmb.Value = valueToSet
    Else
        cmb.AddItem valueToSet
        cmb.Value = valueToSet
    End If
End Sub


' 区切り文字で連結された文字列を元に、リストボックスの項目を選択状態にする


' シートの特定行からデータを読み込み、フォームの各項目に表示する処理
Private Sub LoadDataToForm(ByVal ws As Worksheet, ByVal rowNum As Long)
    m_isComboBoxUpdating = True ' 自動入力中はChangeイベントを無効化
    With ws
        Call SetComboBoxValue(Me.請求書提出先, .Cells(rowNum, "S").Value)
        Me.着手.Value = FormatIfDate(.Cells(rowNum, "G").Value)
        Me.完成.Value = FormatIfDate(.Cells(rowNum, "H").Value)
        Me.金額1.Value = .Cells(rowNum, "K").Value
        Me.担当者.Value = .Cells(rowNum, "C").Value
        Me.提出日付.Value = FormatIfDate(.Cells(rowNum, "Q").Value)
        Call SetTeishutsuyoukou(CStr(.Cells(rowNum, "O").Value))
        
        Call SetDoufuubutsu(CStr(.Cells(rowNum, "P").Value))
        
        Me.引渡日.Value = FormatIfDate(.Cells(rowNum, "R").Value)
        Me.郵便番号.Value = .Cells(rowNum, "T").Value
        Me.住所.Value = .Cells(rowNum, "U").Value
        Me.引継ぎコメント.Value = .Cells(rowNum, "X").Value
    End With
    
    ' 依頼履歴から最新の明細JSONを読み込んでフォームに反映
    Call LoadMeisaiFromHistory(m_KoujiBangou, ws.Parent)
    
    Call CalculateTotals ' 金額を再計算
    m_isComboBoxUpdating = False ' Changeイベントを有効に戻す
End Sub

' 値が日付なら"yyyy/mm/dd"形式の文字列に、そうでなければ空文字に変換する関数
Private Function FormatIfDate(ByVal Value As Variant) As Variant
    If IsDate(Value) Then
        FormatIfDate = Format(CDate(Value), "yyyy/mm/dd")
    Else
        FormatIfDate = ""
    End If
End Function

' このツール内の「工事一覧」シートを、外部ファイルの最新情報に更新する処理
Private Sub UpdateLocalListSheet(ByVal wsSource As Worksheet, ByVal wsMaster As Worksheet)
    If wsSource Is Nothing Or wsMaster Is Nothing Then
        MsgBox "更新元またはマスタシートの参照が不正なため、ローカルシートの更新を中断しました。", vbCritical, "引数エラー"
        Exit Sub
    End If
    Dim wsDest As Worksheet, destSheetName As String, lastRowSource As Long
    Dim copyRange As Range
    ' Dim dataArray As Variant, c As Long, currentColumnWidth As Double ' ← 不要になるためコメントアウトまたは削除
    Dim originalDisplayAlerts As Boolean, originalEnableEvents As Boolean
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    On Error GoTo ErrorHandlerUpdateLocal
    destSheetName = Trim(CStr(wsMaster.Range(CELL_LOCAL_COPY_SHEET).Value))
    If destSheetName = "" Then
        MsgBox "「" & SHEET_KANRI_MASTER & "」" & CELL_LOCAL_COPY_SHEET & "セルにシート名が指定されていません。", vbExclamation
        GoTo FinalizeUpdateLocal
    End If
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(destSheetName)
    On Error GoTo ErrorHandlerUpdateLocal
    If wsDest Is Nothing Then
        MsgBox "このファイルに「" & destSheetName & "」が見つかりませんでした。", vbExclamation
        GoTo FinalizeUpdateLocal
    End If

    ' --- データのコピー処理 ---
    Call SafeUnprotect(wsDest)
    ' .ClearContents から .Clear に変更し、古い書式もクリアする
    wsDest.Range("A3:X" & wsDest.Rows.count).Clear
    
    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    If lastRowSource >= 5 Then
        Set copyRange = wsSource.Range("A5:X" & lastRowSource)
        
        '--- 新しい「書式ごと」コピーする処理 ---
        copyRange.Copy Destination:=wsDest.Range("A3")
        
    End If
    Call SafeProtectData(wsDest)
FinalizeUpdateLocal: ' 終了処理
    Application.CutCopyMode = False
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub
ErrorHandlerUpdateLocal: ' エラー発生時の処理
    MsgBox "ローカルシート更新中に予期せぬエラー発生: " & Err.Description, vbCritical
    Resume FinalizeUpdateLocal
End Sub


' このツール内の「依頼履歴」シートを、外部ファイルの最新情報に更新する処理
Private Sub UpdateLocalRirekiSheet(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet)
    If wsSource Is Nothing Or wsDest Is Nothing Then
        MsgBox "依頼履歴の更新元または更新先シートの参照が不正なため、ローカルシートの更新を中断しました。", vbCritical, "引数エラー"
        Exit Sub
    End If
    
    Dim lastRowSource As Long
    Dim copyRange As Range
    Dim originalDisplayAlerts As Boolean, originalEnableEvents As Boolean
    
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandlerUpdateLocalRireki
    
    ' --- データのコピー処理 ---
    Call SafeUnprotect(wsDest)
    
    ' 3行目以降（ヘッダー(2行目)を除く）のデータをクリア（書式ごと）
    wsDest.Range("A3:W" & wsDest.Rows.count).Clear
    
    ' 外部ファイルのデータ最終行を取得（1行目はヘッダーと仮定）
    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    
    ' 外部ファイルにデータが2行目以降にある場合のみコピー
    If lastRowSource >= 2 Then
        ' コピー範囲はA2からW列の最終行まで（列構成に合わせて"W"に変更）
        Set copyRange = wsSource.Range("A2:W" & lastRowSource)
        
        ' ローカルシートのA3に「書式ごと」コピーする
        copyRange.Copy Destination:=wsDest.Range("A3")
    End If
    
    Call SafeProtectData(wsDest)
FinalizeUpdateLocalRireki: ' 終了処理
    Application.CutCopyMode = False
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub
    
ErrorHandlerUpdateLocalRireki: ' エラー発生時の処理
    MsgBox "ローカルの依頼履歴シート更新中に予期せぬエラー発生: " & Err.Description, vbCritical
    Resume FinalizeUpdateLocalRireki
End Sub



' 提出要項のOptionButtonから選択値を取得する
Private Function GetSelectedTeishutsuyoukou() As String
    If Me.opt提出_普通郵便.Value = True Then
        GetSelectedTeishutsuyoukou = "普通郵便"
    ElseIf Me.opt提出_速達.Value = True Then
        GetSelectedTeishutsuyoukou = "速達"
    ElseIf Me.opt提出_返送.Value = True Then
        GetSelectedTeishutsuyoukou = "依頼者に返送"
    ElseIf Me.opt提出_その他.Value = True Then
        GetSelectedTeishutsuyoukou = "その他"
    Else
        GetSelectedTeishutsuyoukou = ""
    End If
End Function

' 同封のOptionButtonから選択値を取得する
Private Function GetSelectedDoufuubutsu() As String
    If Me.opt同封_見積書.Value = True Then
        GetSelectedDoufuubutsu = "見積書"
    ElseIf Me.opt同封_工事写真.Value = True Then
        GetSelectedDoufuubutsu = "工事写真"
    ElseIf Me.opt同封_完了届.Value = True Then
        GetSelectedDoufuubutsu = "工事完了届"
    ElseIf Me.opt同封_その他.Value = True Then
        If Trim(Me.txt同封その他.Value) <> "" Then
            GetSelectedDoufuubutsu = Me.txt同封その他.Value
        Else
            GetSelectedDoufuubutsu = "その他"
        End If
    Else
        GetSelectedDoufuubutsu = ""
    End If
End Function

' 提出要項の値に応じてOptionButtonを選択状態にする
Private Sub SetTeishutsuyoukou(ByVal val As String)
    Select Case Trim(val)
        Case "普通郵便": Me.opt提出_普通郵便.Value = True
        Case "速達": Me.opt提出_速達.Value = True
        Case "依頼者に返送": Me.opt提出_返送.Value = True
        Case "その他": Me.opt提出_その他.Value = True
        Case Else
            If val <> "" Then Me.opt提出_その他.Value = True
    End Select
End Sub

' 同封物の値に応じてOptionButtonを選択状態にする
Private Sub SetDoufuubutsu(ByVal val As String)
    Select Case Trim(val)
        Case "見積書": Me.opt同封_見積書.Value = True
        Case "工事写真": Me.opt同封_工事写真.Value = True
        Case "工事完了届": Me.opt同封_完了届.Value = True
        Case Else
            If val <> "" Then
                Me.opt同封_その他.Value = True
                Me.txt同封その他.Value = val
            End If
    End Select
End Sub


'================================================================================
' 明細JSON（5行分の名称・数量・単位・金額）シリアライズ／パース
'================================================================================

' フォームの5行分の明細をJSON文字列に変換する
' 空行はスキップ。名称・数量・金額のいずれかが入っていれば書き出す
Private Function BuildMeisaiJSON() As String
    Dim items(1 To 5, 1 To 4) As String  ' (行, フィールド: 名称/数量/単位/金額)
    Dim i As Long
    Dim result As String
    Dim hasAny As Boolean
    
    ' フォームから値を集める
    items(1, 1) = Me.txt名称1.Value: items(1, 2) = Me.数量1.Value: items(1, 3) = Me.単位1.Value: items(1, 4) = Me.金額1.Value
    items(2, 1) = Me.txt名称2.Value: items(2, 2) = Me.数量2.Value: items(2, 3) = Me.単位2.Value: items(2, 4) = Me.金額2.Value
    items(3, 1) = Me.txt名称3.Value: items(3, 2) = Me.数量3.Value: items(3, 3) = Me.単位3.Value: items(3, 4) = Me.金額3.Value
    items(4, 1) = Me.txt名称4.Value: items(4, 2) = Me.数量4.Value: items(4, 3) = Me.単位4.Value: items(4, 4) = Me.金額4.Value
    items(5, 1) = Me.txt名称5.Value: items(5, 2) = Me.数量5.Value: items(5, 3) = Me.単位5.Value: items(5, 4) = Me.金額5.Value
    
    result = "["
    For i = 1 To 5
        ' 空行スキップ判定：名称・数量・金額のどれかが入っていれば採用
        If Trim(items(i, 1)) <> "" Or Trim(items(i, 2)) <> "" Or Trim(items(i, 4)) <> "" Then
            If hasAny Then result = result & ","
            result = result & "{"
            result = result & """name"":""" & JSONEscape(items(i, 1)) & ""","
            result = result & """qty"":" & GetNumOrZero(items(i, 2)) & ","
            result = result & """unit"":""" & JSONEscape(items(i, 3)) & ""","
            result = result & """amount"":" & GetNumOrZero(items(i, 4))
            result = result & "}"
            hasAny = True
        End If
    Next i
    result = result & "]"
    
    If hasAny Then BuildMeisaiJSON = result Else BuildMeisaiJSON = ""
End Function

' JSONエスケープ（"と\のみ対応、通常の工事名では十分）
Private Function JSONEscape(ByVal s As String) As String
    Dim r As String
    r = s
    r = Replace(r, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, vbCrLf, " ")
    r = Replace(r, vbCr, " ")
    r = Replace(r, vbLf, " ")
    JSONEscape = r
End Function

' JSONデコード（エスケープ解除）
Private Function JSONUnescape(ByVal s As String) As String
    Dim r As String
    r = s
    r = Replace(r, "\""", """")
    r = Replace(r, "\\", "\")
    JSONUnescape = r
End Function

' カンマ区切りや単位付きを除いた数値を返す
Private Function GetNumOrZero(ByVal v As Variant) As String
    Dim cleaned As String
    cleaned = Replace(CStr(v), ",", "")
    If IsNumeric(cleaned) And Trim(cleaned) <> "" Then
        GetNumOrZero = cleaned
    Else
        GetNumOrZero = "0"
    End If
End Function

' 依頼履歴シートから該当工事番号の最新明細JSONを読み込み、フォームに反映
Private Sub LoadMeisaiFromHistory(ByVal koujiBangou As String, ByVal wbMaster As Workbook)
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim latestRow As Long
    Dim jsonText As String
    
    On Error GoTo CleanUp
    
    If Trim(koujiBangou) = "" Then Exit Sub
    If wbMaster Is Nothing Then Exit Sub
    
    ' 既に開いている管理表（共有ファイル）の依頼履歴シートを参照
    Set ws = Nothing
    On Error Resume Next
    Set ws = wbMaster.Sheets(SHEET_IRAI_RIREKI)
    On Error GoTo CleanUp
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    latestRow = 0
    
    ' E列(工事番号) でマッチする行を探し、一番下の（最新の）行を取得
    For r = lastRow To 2 Step -1  ' 管理表はヘッダー1行、データは2行目から
        If Trim(CStr(ws.Cells(r, "E").Value)) = Trim(koujiBangou) Then
            latestRow = r
            Exit For
        End If
    Next r
    
    If latestRow = 0 Then Exit Sub
    
    ' Z列から明細JSONを取得
    jsonText = CStr(ws.Cells(latestRow, "Z").Value)
    If Trim(jsonText) = "" Then Exit Sub
    
    ' パースしてフォームに反映
    Call ParseMeisaiJSONToForm(jsonText)
    
CleanUp:
    Exit Sub
End Sub

' JSON文字列をパースして5行分のフォームに設定する
Private Sub ParseMeisaiJSONToForm(ByVal jsonText As String)
    Dim items() As String
    Dim i As Long, rowIdx As Long
    
    ' まず全行をクリア
    Me.txt名称1.Value = "": Me.数量1.Value = "": Me.単位1.Value = "": Me.金額1.Value = ""
    Me.txt名称2.Value = "": Me.数量2.Value = "": Me.単位2.Value = "": Me.金額2.Value = ""
    Me.txt名称3.Value = "": Me.数量3.Value = "": Me.単位3.Value = "": Me.金額3.Value = ""
    Me.txt名称4.Value = "": Me.数量4.Value = "": Me.単位4.Value = "": Me.金額4.Value = ""
    Me.txt名称5.Value = "": Me.数量5.Value = "": Me.単位5.Value = "": Me.金額5.Value = ""
    
    ' オブジェクト単位に分割（}, で区切る）
    Dim working As String
    working = jsonText
    working = Replace(working, "[", "")
    working = Replace(working, "]", "")
    
    ' "},{" で区切ってオブジェクトごとに処理
    items = Split(working, "},{")
    
    rowIdx = 1
    For i = 0 To UBound(items)
        If rowIdx > 5 Then Exit For
        
        Dim obj As String
        obj = items(i)
        ' 両端の { } を除去
        obj = Replace(obj, "{", "")
        obj = Replace(obj, "}", "")
        
        Dim nameVal As String, qtyVal As String, unitVal As String, amountVal As String
        nameVal = ExtractJSONField(obj, "name")
        qtyVal = ExtractJSONField(obj, "qty")
        unitVal = ExtractJSONField(obj, "unit")
        amountVal = ExtractJSONField(obj, "amount")
        
        Select Case rowIdx
            Case 1: Me.txt名称1.Value = nameVal: Me.数量1.Value = qtyVal: Me.単位1.Value = unitVal: Me.金額1.Value = amountVal
            Case 2: Me.txt名称2.Value = nameVal: Me.数量2.Value = qtyVal: Me.単位2.Value = unitVal: Me.金額2.Value = amountVal
            Case 3: Me.txt名称3.Value = nameVal: Me.数量3.Value = qtyVal: Me.単位3.Value = unitVal: Me.金額3.Value = amountVal
            Case 4: Me.txt名称4.Value = nameVal: Me.数量4.Value = qtyVal: Me.単位4.Value = unitVal: Me.金額4.Value = amountVal
            Case 5: Me.txt名称5.Value = nameVal: Me.数量5.Value = qtyVal: Me.単位5.Value = unitVal: Me.金額5.Value = amountVal
        End Select
        
        rowIdx = rowIdx + 1
    Next i
End Sub

' JSONオブジェクト文字列から特定フィールドの値を取り出す
' obj例: "name":"abc","qty":1,"unit":"式","amount":10000
Private Function ExtractJSONField(ByVal obj As String, ByVal fieldName As String) As String
    Dim key As String
    Dim keyPos As Long, valStart As Long, valEnd As Long
    Dim firstChar As String
    
    ExtractJSONField = ""
    
    key = """" & fieldName & """:"
    keyPos = InStr(obj, key)
    If keyPos = 0 Then Exit Function
    
    valStart = keyPos + Len(key)
    If valStart > Len(obj) Then Exit Function
    
    firstChar = Mid(obj, valStart, 1)
    
    If firstChar = """" Then
        ' 文字列値 ："xxx" の形式
        valStart = valStart + 1
        ' 次の " を探す（エスケープされたものは除く）
        valEnd = valStart
        Do While valEnd <= Len(obj)
            If Mid(obj, valEnd, 1) = """" Then
                ' 直前が \ でなければ終端
                If valEnd = 1 Then Exit Do
                If Mid(obj, valEnd - 1, 1) <> "\" Then Exit Do
            End If
            valEnd = valEnd + 1
        Loop
        ExtractJSONField = JSONUnescape(Mid(obj, valStart, valEnd - valStart))
    Else
        ' 数値値：次のカンマか末尾まで
        valEnd = InStr(valStart, obj, ",")
        If valEnd = 0 Then valEnd = Len(obj) + 1
        ExtractJSONField = Trim(Mid(obj, valStart, valEnd - valStart))
    End If
End Function



'================================================================================
' セル位置マッピング（依頼書セル設定シートから動的に読み込む）
'================================================================================

' 「依頼書セル設定」シートからセル位置マッピングを読み込んでキャッシュ
Private Sub LoadCellSettings(ByVal wbMaster As Workbook)
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim itemName As String, cellAddr As String
    
    Set m_CellMap = CreateObject("Scripting.Dictionary")
    
    On Error GoTo CleanUp
    If wbMaster Is Nothing Then Exit Sub
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = wbMaster.Sheets(SHEET_CELL_SETTING)
    On Error GoTo CleanUp
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For r = 2 To lastRow
        itemName = Trim(CStr(ws.Cells(r, "A").Value))
        cellAddr = Trim(CStr(ws.Cells(r, "B").Value))
        ' セクション見出し（【...】）やセルが空の行はスキップ
        If itemName <> "" And cellAddr <> "" And Left(itemName, 1) <> "【" Then
            If Not m_CellMap.Exists(itemName) Then
                m_CellMap.Add itemName, cellAddr
            End If
        End If
    Next r
    
CleanUp:
    Exit Sub
End Sub

' 項目名からセル番地を取得（単一セル用）
Private Function GetCellAddr(ByVal itemName As String) As String
    If m_CellMap Is Nothing Then GetCellAddr = "": Exit Function
    If m_CellMap.Exists(itemName) Then
        GetCellAddr = m_CellMap(itemName)
    Else
        GetCellAddr = ""
    End If
End Function

' 明細セルの番地を取得（rowIdx: 1～5）
' 例: GetMeisaiCellAddr("名称列", 1) → "F14"
Private Function GetMeisaiCellAddr(ByVal columnItemName As String, ByVal rowIdx As Long) As String
    Dim col As String, startRow As String, r As Long
    col = GetCellAddr(columnItemName)
    startRow = GetCellAddr("明細開始行")
    If col = "" Or startRow = "" Or Not IsNumeric(startRow) Then
        GetMeisaiCellAddr = ""
        Exit Function
    End If
    r = CLng(startRow) + (rowIdx - 1)
    GetMeisaiCellAddr = col & r
End Function

' 小計セル番地：明細開始行 + 明細行数 行目 + 小計列
Private Function GetSubtotalCellAddr() As String
    Dim col As String, startRow As String
    col = GetCellAddr("小計列")
    startRow = GetCellAddr("明細開始行")
    If col = "" Or startRow = "" Or Not IsNumeric(startRow) Then
        GetSubtotalCellAddr = ""
        Exit Function
    End If
    GetSubtotalCellAddr = col & (CLng(startRow) + MEISAI_ROW_COUNT)
End Function

' 消費税セル番地：小計の次の行 + 消費税列
Private Function GetTaxCellAddr() As String
    Dim col As String, startRow As String
    col = GetCellAddr("消費税列")
    startRow = GetCellAddr("明細開始行")
    If col = "" Or startRow = "" Or Not IsNumeric(startRow) Then
        GetTaxCellAddr = ""
        Exit Function
    End If
    GetTaxCellAddr = col & (CLng(startRow) + MEISAI_ROW_COUNT + 1)
End Function

' 明細1行を転記するヘルパー
Private Sub WriteMeisaiRow(ByVal ws As Worksheet, ByVal rowIdx As Long, _
                            ByVal nameVal As String, ByVal qtyVal As String, _
                            ByVal unitVal As String, ByVal amountVal As String)
    ws.Range(GetMeisaiCellAddr("名称列", rowIdx)).Value = nameVal
    ws.Range(GetMeisaiCellAddr("数量列", rowIdx)).Value = val(qtyVal)
    ws.Range(GetMeisaiCellAddr("単位列", rowIdx)).Value = unitVal
    ws.Range(GetMeisaiCellAddr("単価列", rowIdx)).Value = val(Replace(amountVal, ",", ""))
    ws.Range(GetMeisaiCellAddr("金額列", rowIdx)).Value = val(Replace(amountVal, ",", ""))
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
