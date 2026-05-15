Option Explicit

'=========================================================
' 設定エリア：転記先のセル番地などを一括管理
'=========================================================
'--- 参照元（依頼検索シート）のセル ---
Private Const SRC_KEIRI_NO As String = "B2"       ' 請求書番号(固定)
Private Const SRC_PROJECT_NO As String = "A5"     ' 工事番号
Private Const SRC_DATE_SUBMISSION As String = "B5" ' 提出日付
Private Const SRC_DESTINATION As String = "C5"    ' 請求書提出先

' Main_Search_Updateに合わせて修正 (D5:名称, E5:金額)
Private Const SRC_PROJECT_NAME As String = "D5"   ' 工事名称
Private Const SRC_AMOUNT As String = "E5"         ' 請負金額
Private Const SRC_STAFF_NAME As String = "F5"     ' 担当者

Private Const SRC_START_DATE As String = "A7"     ' 工期着手
Private Const SRC_END_DATE As String = "B7"       ' 工期完成

'--- 転記先（請求書シート）のセル（変更なし） ---
Private Const DEST_KEIRI_NO As String = "AF1"
Private Const DEST_DATE_YEAR As String = "AB3"
Private Const DEST_DATE_MONTH As String = "AE3"
Private Const DEST_DATE_DAY As String = "AG3"
Private Const DEST_DESTINATION As String = "C5:N5"
Private Const DEST_HONORIFIC As String = "O5"
Private Const DEST_DEPT_NAME As String = "AB15:AD15"
Private Const DEST_STAFF_NAME As String = "AG15:AH15"
Private Const DEST_PROJECT_NAME As String = "C18:T18"
Private Const DEST_SUBJECT As String = "F19"
Private Const DEST_AMOUNT As String = "AA19:AE19"
Private Const DEST_PROJECT_NO As String = "AK2"
Private Const DEST_FILENAME_CELL As String = "AK1"

'--- 工期転記先（日付分解用） ---
Private Const DEST_TERM_YEAR As String = "A18"
Private Const DEST_TERM_START_MONTH As String = "A19"
Private Const DEST_TERM_START_DAY As String = "B19"
Private Const DEST_TERM_TILDE As String = "B20"
Private Const DEST_TERM_END_YEAR As String = "A20"
Private Const DEST_TERM_END_MONTH As String = "A21"
Private Const DEST_TERM_END_DAY As String = "B21"

'=========================================================
' 【メイン】請求書作成メニュー
'=========================================================
Sub 請求書作成メニュー()
    Dim wsMenu As Worksheet
    Dim chkBox As CheckBox
    Dim btn As Button
    Dim i As Integer
    Dim arrLabels As Variant
    Dim topPos As Double
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("作成選択").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsMenu = Worksheets.Add
    wsMenu.Name = "作成選択"
    
    With wsMenu.Range("B2")
        .Value = "作成する請求書にチェックを入れてください"
        .Font.Bold = True
        .Font.Size = 14
    End With
    
    ' ★修正：カッコを全て半角に統一
    ' ※Excel側のシート名も全て半角カッコ「(」「)」になっているか確認してください。
    arrLabels = Array("請求書", _
                      "請求書(振込手数料有)", _
                      "請求書(領収有)", _
                      "請求書(領収有・振込手数料有)", _
                      "但陽信用金庫")
    
    topPos = 60
    For i = 0 To UBound(arrLabels)
        Set chkBox = wsMenu.CheckBoxes.Add(Left:=50, Top:=topPos, Width:=250, Height:=20)
        With chkBox
            .Caption = arrLabels(i)
            .Name = "chkOption_" & (i + 1)
            .Value = xlOff
        End With
        topPos = topPos + 30
    Next i
    
    Set btn = wsMenu.Buttons.Add(Left:=50, Top:=topPos + 20, Width:=150, Height:=40)
    With btn
        .Caption = "選択した請求書を作成"
        .Font.Bold = True
        .Font.Size = 12
        .OnAction = "'GenerateFromSelection'"
    End With
    
    ActiveWindow.DisplayGridlines = False
    wsMenu.Range("A1").Select
    Application.ScreenUpdating = True
End Sub

Sub GenerateFromSelection()
    Dim wsMenu As Worksheet
    Dim wsSearch As Worksheet: Set wsSearch = ThisWorkbook.Worksheets("依頼検索")
    Dim i As Integer
    Dim anySelected As Boolean
    Dim isSelected(1 To 5) As Boolean
    Dim honorificValue As String
    
    On Error Resume Next
    Set wsMenu = ThisWorkbook.Worksheets("作成選択")
    On Error GoTo 0
    
    If wsMenu Is Nothing Then
        Exit Sub
    End If
    
    anySelected = False
    For i = 1 To 5
        If wsMenu.CheckBoxes("chkOption_" & i).Value = xlOn Then
            isSelected(i) = True
            anySelected = True
        End If
    Next i
    
    If Not anySelected Then
        MsgBox "請求書が選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' 敬称確認（1回のみ実行）
    honorificValue = DetermineHonorific(Trim(wsSearch.Range(SRC_DESTINATION).Value))
    
    Call AssignAccountingInvoiceNo
    
    Application.DisplayAlerts = False
    wsMenu.Delete
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = False
    
    ' 文字列を統一して呼び出し
    If isSelected(1) Then Call CreateInvoiceFromTemplate("請求書", honorificValue)
    If isSelected(2) Then Call CreateInvoiceFromTemplate("請求書(振込手数料有)", honorificValue)
    If isSelected(3) Then Call CreateInvoiceFromTemplate("請求書(領収有)", honorificValue)
    If isSelected(4) Then Call CreateInvoiceFromTemplate("請求書(領収有・振込手数料有)", honorificValue)
    If isSelected(5) Then Call CreateInvoiceFromTemplate("但陽信用金庫", honorificValue)
    
    Application.ScreenUpdating = True
    MsgBox "請求書の作成が完了しました。", vbInformation
End Sub

Public Sub AssignAccountingInvoiceNo()
    Dim wsRireki As Worksheet: Set wsRireki = ThisWorkbook.Sheets("依頼履歴")
    Dim wsSearch As Worksheet: Set wsSearch = ThisWorkbook.Sheets("依頼検索")
    Dim lastRow As Long, maxNo As Double, f As Range
    Dim irNo As String: irNo = Trim(wsSearch.Range("A2").Value) ' 検索キーはA2固定

    If irNo = "" Then Exit Sub
    ' 経理番号が空なら採番
    If wsSearch.Range(SRC_KEIRI_NO).Value <> "" Then Exit Sub

    lastRow = wsRireki.Cells(wsRireki.Rows.Count, "B").End(xlUp).Row
    If lastRow < 3 Then
        maxNo = 0
    Else
        maxNo = Application.WorksheetFunction.Max(wsRireki.Range("B3:B" & lastRow))
    End If
    
    wsSearch.Range(SRC_KEIRI_NO).Value = Int(maxNo) + 1
    Set f = wsRireki.Columns("A").Find(What:=irNo, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        wsRireki.Cells(f.Row, "B").Value = wsSearch.Range(SRC_KEIRI_NO).Value
    End If
End Sub

Private Sub CreateInvoiceFromTemplate(templateName As String, honorific As String)
    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets("依頼検索")
    Dim newSheetName As String
    Dim templateWs As Worksheet
    Dim targetWs As Worksheet ' 作成されるシート（コピー先）
    
    ' ブック保護チェック
    If ThisWorkbook.ProtectStructure Then
        MsgBox "ブックの構成が保護されているため、シートをコピーできません。" & vbCrLf & _
               "Excelの「校閲」タブから「ブックの保護」を解除してから実行してください。", vbCritical
        Exit Sub
    End If
    
    ' テンプレート取得チェック（ループチェックに変更）
    If Not SheetExists(templateName) Then
        MsgBox "エラー：テンプレートシート「" & templateName & "」が見つかりません。", vbCritical
        Exit Sub
    End If
    Set templateWs = ThisWorkbook.Sheets(templateName)
    
    ' 非表示対策
    Dim originalVisibleState As XlSheetVisibility
    originalVisibleState = templateWs.Visible
    If originalVisibleState <> xlSheetVisible Then templateWs.Visible = xlSheetVisible
    
    ' -----------------------------------------------------
    ' 【最強のコピー処理】
    ' -----------------------------------------------------
    On Error GoTo FallbackCopyMode
    
    ' 1. 通常のコピーを試行
    templateWs.Copy After:=wsSrc
    Set targetWs = ActiveSheet
    
    On Error GoTo 0 ' エラー解除
    GoTo PostCopyProcess
    
FallbackCopyMode:
    ' 2. 通常コピー失敗時の救済措置（手動コピーモード）
    Err.Clear
    On Error Resume Next
    
    ' 空のシートを追加
    Set targetWs = Worksheets.Add(After:=wsSrc)
    
    ' 中身を全てコピーして貼り付け
    templateWs.Cells.Copy targetWs.Range("A1")
    
    ' 【重要】PageSetup関連のエラー回避のため、コードを無効化しています。
    ' 印刷設定が必要な場合は手動で設定するか、ドライバの問題を解決後に有効化してください。
    ' targetWs.PageSetup.PrintArea = templateWs.PageSetup.PrintArea
    ' targetWs.PageSetup.Orientation = templateWs.PageSetup.Orientation
    
    On Error GoTo 0

PostCopyProcess:
    ' 元のシートの表示状態を戻す
    If Not templateWs Is Nothing Then
        If originalVisibleState <> xlSheetVisible Then templateWs.Visible = originalVisibleState
    End If

    ' ---------------------------------------------------------
    ' データ転記処理（targetWs に対して行う）
    ' ---------------------------------------------------------
    If targetWs Is Nothing Then
        MsgBox "シートの作成に失敗しました。", vbCritical
        Exit Sub
    End If
    
    targetWs.Activate
    Call TransferData(wsSrc, targetWs, honorific)
    
    newSheetName = "No_" & targetWs.Range(DEST_KEIRI_NO).Value & "_" & templateName
    If Len(newSheetName) > 31 Then newSheetName = Left(newSheetName, 31)
    
    ' ★修正：同名シート削除処理の強化（エラー9回避）
    ' On Error Resume Nextに頼らず、存在確認関数(SheetExists)を使って安全に削除します
    If SheetExists(newSheetName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(newSheetName).Delete
        Application.DisplayAlerts = True
    End If
    
    ' シート名を変更（エラーハンドリング付き）
    On Error Resume Next
    targetWs.Name = newSheetName
    If Err.Number <> 0 Then
        MsgBox "シート名の変更に失敗しました。" & vbCrLf & "名前：" & newSheetName & vbCrLf & _
               "原因：同名のシートが残っているか、名前に使用できない文字が含まれています。", vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
    
    targetWs.Range("A18").Select
End Sub

' ★修正：エラー設定に依存しない、ループ方式の存在チェック
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    SheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If LCase(ws.Name) = LCase(sheetName) Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Private Sub TransferData(wsSource As Worksheet, wsDest As Worksheet, honorific As String)
    wsDest.Range(DEST_KEIRI_NO).Value = wsSource.Range(SRC_KEIRI_NO).Value
    
    Dim dSub As Variant: dSub = wsSource.Range(SRC_DATE_SUBMISSION).Value
    ' 提出日付が空や非日付の場合は今日の日付をフォールバックとして使用
    Dim useDate As Date
    If IsDate(dSub) Then
        useDate = CDate(dSub)
    Else
        useDate = Date
    End If
    wsDest.Range(DEST_DATE_YEAR).Value = Year(useDate)
    wsDest.Range(DEST_DATE_MONTH).Value = Month(useDate)
    wsDest.Range(DEST_DATE_DAY).Value = Day(useDate)
    
    Dim dName As String: dName = Trim(wsSource.Range(SRC_DESTINATION).Value)
    wsDest.Range(DEST_DESTINATION).Value = dName
    wsDest.Range(DEST_HONORIFIC).Value = honorific

    Dim pNo As String: pNo = Trim(wsSource.Range(SRC_PROJECT_NO).Value)
    Dim pName As String: pName = Trim(wsSource.Range(SRC_PROJECT_NAME).Value)
    
    Dim deptV: On Error Resume Next
    deptV = Application.VLookup(Left(pNo, 2), ThisWorkbook.Sheets("部署マスタ").Range("A:B"), 2, False)
    On Error GoTo 0
    wsDest.Range(DEST_DEPT_NAME).Value = IIf(IsError(deptV) Or deptV = "", "建築事業部", deptV)
    
    wsDest.Range(DEST_STAFF_NAME).Value = wsSource.Range(SRC_STAFF_NAME).Value
    wsDest.Range(DEST_PROJECT_NAME).Value = pName
    wsDest.Range(DEST_SUBJECT).Value = "上記工事代"
    
    If IsNumeric(wsSource.Range(SRC_AMOUNT).Value) Then
        wsDest.Range(DEST_AMOUNT).Value = Int(wsSource.Range(SRC_AMOUNT).Value / 1.1)
    End If
    wsDest.Range(DEST_PROJECT_NO).Value = pNo
    
    '--- 工期転記 ---
    Dim dStart As Variant: dStart = wsSource.Range(SRC_START_DATE).Value
    Dim dEnd As Variant: dEnd = wsSource.Range(SRC_END_DATE).Value
    
    If IsDate(dStart) Then
        wsDest.Range(DEST_TERM_YEAR).Value = Year(dStart)
        wsDest.Range(DEST_TERM_START_MONTH).Value = Month(dStart)
        wsDest.Range(DEST_TERM_START_DAY).Value = Day(dStart)
    End If
    wsDest.Range(DEST_TERM_TILDE).Value = "～"
    If IsDate(dEnd) Then
        wsDest.Range(DEST_TERM_END_MONTH).Value = Month(dEnd)
        wsDest.Range(DEST_TERM_END_DAY).Value = Day(dEnd)
        If IsDate(dStart) Then
            If Year(dStart) <> Year(dEnd) Then
                wsDest.Range(DEST_TERM_END_YEAR).Value = Year(dEnd)
            End If
        End If
    End If
    
    ' --- ★宛名の文字サイズ・折り返し自動調整 ---
    Dim dr As Range: Set dr = wsDest.Range(DEST_DESTINATION)
    Dim valLen As Long: valLen = Len(dr.Cells(1).Value)
    
    dr.ShrinkToFit = False ' 一旦リセット
    dr.WrapText = False    ' 一旦リセット
    
    Select Case valLen
        Case Is <= 13
            ' 短い場合：大きな文字で1行
            dr.Font.Size = 20
            dr.ShrinkToFit = True
            
        Case Is <= 25
            ' 少し長い場合：中くらいの文字で1行（縮小あり）
            dr.Font.Size = 14
            dr.ShrinkToFit = True
            
        Case Else
            ' かなり長い場合：読めるサイズ(11pt)にして、2行に折り返す
            dr.Font.Size = 11
            dr.WrapText = True
    End Select
    ' ------------------------------------------
    
    Dim fn As String: fn = wsSource.Range(SRC_KEIRI_NO).Value & "_請求書_" & dName & "_" & pName & "_" & pNo
    Dim invalid, c: invalid = Array("/", "\", ":", "*", "?", "<", ">", "|", "[", "]")
    For Each c In invalid: fn = Replace(fn, c, " "): Next
    wsDest.Range(DEST_FILENAME_CELL).Value = Trim(fn)
End Sub

Private Function DetermineHonorific(ByVal t As String) As String
    If MsgBox("宛名：「" & t & "」の敬称は『様』でよろしいですか？", vbYesNo + vbQuestion) = vbNo Then
        DetermineHonorific = "御中"
    Else
        DetermineHonorific = "様"
    End If
End Function

