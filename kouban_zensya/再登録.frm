VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 再登録 
   Caption         =   "再登録"
   ClientHeight    =   9132
   ClientLeft      =   96
   ClientTop       =   360
   ClientWidth     =   9000
   OleObjectBlob   =   "再登録.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "再登録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ユーザーフォーム: 再登録
' IS_TEST_MODE / TEST_FILE_PATH / SHEET_* / CELL_* / PATH_CELL は Config モジュールで一元管理。

Option Explicit

'================================================================================
' [フォーム] 再登録（既存の工事データを編集するフォーム）
'--------------------------------------------------------------------------------
' 【何のファイル？】
'   工事名称＋担当者の組で既存行を検索し、フォーム上で編集して
'   マスタに上書き保存する処理を持つフォーム。
'
' 【しくみ】
'   - 担当者リストは担当者マスタから自部署フィルタ
'   - 工事番号一覧シート名は管理マスタの key/value から取得
'================================================================================


'--- Public 変数（他フォームから受け取る） ---
Public SearchedKoujiName As String
Public SelectedTantousha As String

'--- Private 変数 ---
Private Const COL_STAFF As String = "C"
Private Const COL_KOUJI_NAME As String = "E"
Private Const MASTER_COL_STAFF_NAME As String = "A"

Private m_CachedStaffList As Variant
Private m_CachedTargetSheetName As String
Private m_TargetRow As Long

Private Sub Label18_Click()
End Sub

'================================================================================
' フォーム初期化
'================================================================================

Private Sub UserForm_Terminate()
    ' メモリリーク防止：キャッシュ変数を明示的に解放
    m_CachedStaffList = Empty
End Sub

Private Sub UserForm_Initialize()
    Dim wbTarget_Init As Workbook
    Dim wsMaster_Init As Worksheet
    Dim targetFilePath As String
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean

    Application.ScreenUpdating = False
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    originalEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    m_TargetRow = 0
    Me.再登録.Enabled = False

    On Error GoTo ErrorHandlerInit

    If Not IsEmpty(m_CachedStaffList) And m_CachedTargetSheetName <> "" Then
        Me.担当者.List = m_CachedStaffList
        GoTo FinalizeInitWithoutFileOpen
    End If

    targetFilePath = GetTargetFilePath()

    If Dir(targetFilePath) = "" Then
        MsgBox "対象ファイルが見つかりません。", vbCritical
        Unload Me
        GoTo FinalizeInit
    End If

    Me.担当者.Clear
    Set wbTarget_Init = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)

    ' 工事番号一覧シート名を管理マスタから取得
    m_CachedTargetSheetName = GetSheetNameFromMaster(wbTarget_Init, "工事番号一覧シート", SHEET_KOUJI_LIST)
    If m_CachedTargetSheetName = "" Then
        MsgBox "「管理マスタ」シートの「工事番号一覧シート」項目に値が設定されていません。", vbCritical, "設定エラー"
        GoTo FinalizeInit
    End If

    If Not SheetExists(wbTarget_Init, m_CachedTargetSheetName) Then
        MsgBox "参照ファイルに「" & m_CachedTargetSheetName & "」が見つかりません。", vbCritical, "シートエラー"
        GoTo FinalizeInit
    End If

    ' 担当者リスト = 担当者マスタから自部門でフィルタ
    m_CachedStaffList = GetStaffListFiltered(wbTarget_Init, GetMyDeptCode())
    If IsEmpty(m_CachedStaffList) Then
        MsgBox "担当者マスタに自部門(" & GetMyDeptCode() & ")の担当者が見つかりません。", vbExclamation
        GoTo FinalizeInit
    End If
    Me.担当者.List = m_CachedStaffList

FinalizeInit:
    If Not wbTarget_Init Is Nothing Then wbTarget_Init.Close SaveChanges:=False

FinalizeInitWithoutFileOpen:
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandlerInit:
    MsgBox "初期化中に予期せぬエラー発生: " & Err.Description, vbCritical
    Resume FinalizeInit
End Sub

'================================================================================
' フォームアクティブ時の処理
'================================================================================
Private Sub UserForm_Activate()
    Dim wbTarget_Activate As Workbook
    Dim wsTarget_Activate As Worksheet
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim targetFilePath As String

    Application.ScreenUpdating = False
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    originalEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerActivate

    If Trim(Me.SearchedKoujiName) = "" Or Trim(Me.SelectedTantousha) = "" Then
        MsgBox "検索する工事名称または担当者が指定されていません。", vbExclamation
        Unload Me
        GoTo FinalizeActivate
    End If

    If IsEmpty(m_CachedStaffList) Or m_CachedTargetSheetName = "" Then
        MsgBox "キャッシュデータが不足しています。フォームを一度閉じて再度開いてください。", vbCritical
        Unload Me
        GoTo FinalizeActivate
    End If

    targetFilePath = GetTargetFilePath()

    If Dir(targetFilePath) = "" Then
        MsgBox "対象ファイルが見つかりません。", vbCritical
        Unload Me
        GoTo FinalizeActivate
    End If

    Set wbTarget_Activate = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)

    If Not SheetExists(wbTarget_Activate, m_CachedTargetSheetName) Then
        MsgBox "外部ファイルにシート「" & m_CachedTargetSheetName & "」が見つかりません。", vbCritical
        GoTo FinalizeActivate
    End If
    Set wsTarget_Activate = wbTarget_Activate.Sheets(m_CachedTargetSheetName)

    Dim foundRow As Long, r As Long
    Dim sheetKoujiName As String, sheetStaffName As String, sheetKoujiBango As String
    foundRow = 0

    ' 部門フィルタ用 prefix
    Dim deptPrefix_Re As String
    On Error Resume Next
    deptPrefix_Re = GetMyDeptCode() & "-"
    On Error GoTo 0
    If deptPrefix_Re = "-" Then deptPrefix_Re = "03-"

    For r = 2 To wsTarget_Activate.Cells(wsTarget_Activate.Rows.count, "E").End(xlUp).Row
        sheetKoujiName = Trim(CStr(wsTarget_Activate.Cells(r, COL_KOUJI_NAME).Value))
        sheetStaffName = Trim(CStr(wsTarget_Activate.Cells(r, COL_STAFF).Value))
        sheetKoujiBango = Trim(CStr(wsTarget_Activate.Cells(r, "D").Value))
        ' 自部門の工事のみマッチ対象（同名工事が他部署にあっても誤マッチ防止）
        If Left(sheetKoujiBango, Len(deptPrefix_Re)) <> deptPrefix_Re Then GoTo NextRow_Re
        If sheetKoujiName = Trim(Me.SearchedKoujiName) And sheetStaffName = Trim(Me.SelectedTantousha) Then
            foundRow = r
            Exit For
        End If
NextRow_Re:
    Next r

    If foundRow = 0 Then
        MsgBox "「" & Me.SelectedTantousha & "」の工事「" & Me.SearchedKoujiName & "」が見つかりませんでした。", vbExclamation
        Unload Me
        GoTo FinalizeActivate
    End If

    m_TargetRow = foundRow
    Call LoadDataToForm(wsTarget_Activate, m_TargetRow)
    Me.再登録.Enabled = True
    Me.工事名称.SetFocus

FinalizeActivate:
    If Not wsTarget_Activate Is Nothing Then Set wsTarget_Activate = Nothing
    If Not wbTarget_Activate Is Nothing Then wbTarget_Activate.Close SaveChanges:=False
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandlerActivate:
    MsgBox "データ読み込み中に予期せぬエラー発生: " & Err.Description, vbCritical
    Resume FinalizeActivate
End Sub

'================================================================================
' 「再登録」ボタン
'================================================================================
Private Sub 再登録_Click()
    Dim errorMessage As String
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim wbTarget_Reg As Workbook
    Dim wsTarget_Reg As Worksheet, wsMaster_Reg As Worksheet
    Dim targetFilePath As String
    Dim isSuccess As Boolean

    isSuccess = False

    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerReg

    If Me.担当者.ListIndex = -1 Then errorMessage = errorMessage & "・担当者名を選択してください。" & vbCrLf
    If Trim(Me.工事名称.Value) = "" Then errorMessage = errorMessage & "・工事名称を入力してください。" & vbCrLf
    If Trim(Me.発注者入力.Value) = "" Then errorMessage = errorMessage & "・発注者を入力してください。" & vbCrLf

    If errorMessage <> "" Then
        MsgBox "以下の項目を確認してください。" & vbCrLf & vbCrLf & errorMessage, vbExclamation, "入力エラー"
        GoTo TheEndReg
    End If

    If m_TargetRow = 0 Then
        MsgBox "更新対象の行が特定されていません。", vbCritical
        GoTo TheEndReg
    End If

    If MsgBox("入力内容で工事情報を上書きします。よろしいですか？", vbQuestion + vbYesNo, "再登録確認") = vbNo Then
        GoTo TheEndReg
    End If

    targetFilePath = GetTargetFilePath()

    Dim openedWbMain As Workbook
    For Each openedWbMain In Application.Workbooks
        If openedWbMain.FullName = targetFilePath Then
            MsgBox "対象のExcelファイルが既に開かれています。閉じてから再度実行してください。", vbCritical
            GoTo TheEndReg
        End If
    Next openedWbMain

    On Error Resume Next
    Set wbTarget_Reg = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=False, UpdateLinks:=0)
    On Error GoTo ErrorHandlerReg

    If wbTarget_Reg Is Nothing Then
        MsgBox "対象のExcelファイルを開けませんでした。他のユーザーが書き込みロックしている可能性があります。", vbCritical
        GoTo TheEndReg
    End If

    If wbTarget_Reg.ReadOnly Then
        MsgBox "対象のExcelファイルは読み取り専用で開かれました。このため、変更を保存できません。", vbExclamation
        GoTo TheEndReg
    End If

    ' 互換用シート参照（実際の値は GetSheetNameFromMaster を使うようになった）
    If Not SheetExists(wbTarget_Reg, SHEET_KANRI_MASTER) Then
        MsgBox "登録先Excelに「" & SHEET_KANRI_MASTER & "」が見つかりません。", vbCritical
        GoTo TheEndReg
    End If
    Set wsMaster_Reg = wbTarget_Reg.Sheets(SHEET_KANRI_MASTER)

    If m_CachedTargetSheetName = "" Then
        MsgBox "対象シート名がキャッシュされていません。フォームを一度閉じて再度開いてください。", vbCritical
        GoTo TheEndReg
    End If

    If Not SheetExists(wbTarget_Reg, m_CachedTargetSheetName) Then
        MsgBox "登録先Excelにシート「" & m_CachedTargetSheetName & "」が見つかりません。", vbCritical
        GoTo TheEndReg
    End If
    Set wsTarget_Reg = wbTarget_Reg.Sheets(m_CachedTargetSheetName)

    Call SafeUnprotect(wsTarget_Reg)

    With wsTarget_Reg
        .Cells(m_TargetRow, "C").Value = Me.担当者.Value
        .Cells(m_TargetRow, "E").Value = Me.工事名称.Value
        .Cells(m_TargetRow, "F").Value = Me.発注者入力.Value
        If IsDate(Me.着手.Value) Then .Cells(m_TargetRow, "G").Value = CDate(Me.着手.Value) Else .Cells(m_TargetRow, "G").Value = Empty
        If IsDate(Me.完成.Value) Then .Cells(m_TargetRow, "H").Value = CDate(Me.完成.Value) Else .Cells(m_TargetRow, "H").Value = Empty
        If Me.契約有り.Value = True Then
            .Cells(m_TargetRow, "I").Value = "◯"
        ElseIf Me.契約無し.Value = True Then
            .Cells(m_TargetRow, "I").Value = "ー"
        Else
            .Cells(m_TargetRow, "I").Value = ""
        End If
        If IsDate(Me.契約日.Value) Then .Cells(m_TargetRow, "J").Value = CDate(Me.契約日.Value) Else .Cells(m_TargetRow, "J").Value = Empty
        .Cells(m_TargetRow, "K").Value = Me.金額.Value
        If Me.アンケート.Value = True Then .Cells(m_TargetRow, "L").Value = "◯" Else .Cells(m_TargetRow, "L").Value = "ー"
        .Cells(m_TargetRow, "M").Value = Me.備考.Value
    End With
    Call SafeProtectData(wsTarget_Reg)

    Call ClearAllFilters(wsTarget_Reg)
    wbTarget_Reg.Save
    Call UpdateLocalListSheet(wsTarget_Reg, wsMaster_Reg)
    isSuccess = True

TheEndReg:
    If Not wbTarget_Reg Is Nothing Then wbTarget_Reg.Close SaveChanges:=False
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    If isSuccess Then
        MsgBox "工事情報を更新しました。", vbInformation
        Unload Me
    End If
    Exit Sub

ErrorHandlerReg:
    MsgBox "予期せぬエラー発生: " & Err.Description, vbCritical
    Resume TheEndReg
End Sub

'================================================================================
' 補助処理
'================================================================================
Private Sub LoadDataToForm(ByVal ws As Worksheet, ByVal rowNum As Long)
    With ws
        Me.担当者.Value = .Cells(rowNum, "C").Value
        Me.工事名称.Value = .Cells(rowNum, "E").Value
        Me.発注者入力.Value = .Cells(rowNum, "F").Value
        Me.着手.Value = FormatIfDate(.Cells(rowNum, "G").Value)
        Me.完成.Value = FormatIfDate(.Cells(rowNum, "H").Value)
        Dim strKeiyakuUmu As String
        strKeiyakuUmu = Trim(CStr(.Cells(rowNum, "I").Value))
        If strKeiyakuUmu = "◯" Then Me.契約有り.Value = True Else Me.契約無し.Value = True
        Me.契約日.Value = FormatIfDate(.Cells(rowNum, "J").Value)
        Me.金額.Value = .Cells(rowNum, "K").Value
        If Trim(CStr(.Cells(rowNum, "L").Value)) = "◯" Then Me.アンケート.Value = True
        Me.備考.Value = .Cells(rowNum, "M").Value
    End With
End Sub

Private Sub 着手_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.着手, "着手日", Cancel): End Sub
Private Sub 完成_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.完成, "完成日", Cancel): End Sub
Private Sub 契約日_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call ValidateDate(Me.契約日, "契約日", Cancel): End Sub

Private Sub ValidateDate(ByVal DateField As MSForms.TextBox, ByVal FieldName As String, ByRef Cancel As MSForms.ReturnBoolean)
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

Private Sub UpdateLocalListSheet(ByVal wsSource As Worksheet, ByVal wsMaster As Worksheet)
    Dim wsDest As Worksheet
    Dim destSheetName As String
    Dim lastRowSource As Long
    Dim copyRange As Range

    On Error GoTo ErrorHandlerUpdateLocal

    destSheetName = GetSheetNameFromMaster(ThisWorkbook, "工事番号一覧シート", SHEET_KOUJI_LIST)
    If destSheetName = "" Then Exit Sub

    Set wsDest = ThisWorkbook.Sheets(destSheetName)
    If wsDest Is Nothing Then Exit Sub

    On Error GoTo ErrorHandlerUpdateLocalSai
    Call SafeUnprotect(wsDest)
    wsDest.Range("A3:X" & wsDest.Rows.count).Clear
    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    If lastRowSource >= 5 Then
        Set copyRange = wsSource.Range("A5:X" & lastRowSource)
        copyRange.Copy Destination:=wsDest.Range("A3")
    End If

    Call SafeProtectData(wsDest)
    Application.CutCopyMode = False
    Exit Sub

ErrorHandlerUpdateLocalSai:
    Call SafeProtectData(wsDest)
    MsgBox "ローカル一覧更新中にエラー: " & Err.Description, vbCritical
    Exit Sub

ErrorHandlerUpdateLocal:
    MsgBox "ローカルシート更新中にエラー発生: " & Err.Description, vbCritical
End Sub


