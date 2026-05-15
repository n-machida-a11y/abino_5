VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 削除フォーム 
   Caption         =   "削除フォーム"
   ClientHeight    =   6696
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8460
   OleObjectBlob   =   "削除フォーム.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "削除フォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ユーザーフォーム: 削除フォーム
' IS_TEST_MODE / TEST_FILE_PATH / SHEET_* / CELL_* / PATH_CELL は Config モジュールで一元管理。

Option Explicit

' 外部ファイルのパスをキャッシュする（Initialize で一度だけ取得）
Private m_TARGET_FILE_PATH As String

'================================================================================
' フォーム初期化
'================================================================================

Private Sub UserForm_Terminate()
    ' メモリリーク防止：Dictionary オブジェクトを明示的に解放
    Set m_CachedKoujiData = Nothing
End Sub

Private Sub UserForm_Initialize()
    Dim wbTarget As Workbook, wsMaster As Worksheet
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean

    Me.削除.Enabled = False

    Application.ScreenUpdating = False
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    originalEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerInit

    m_TARGET_FILE_PATH = GetTargetFilePath()

    If Dir(m_TARGET_FILE_PATH) = "" Then
        MsgBox "指定されたファイルが見つかりません。", vbCritical
        Unload Me
        GoTo FinalizeInit
    End If

    Set wbTarget = Application.Workbooks.Open(fileName:=m_TARGET_FILE_PATH, ReadOnly:=True, UpdateLinks:=0)

    If Not SheetExists(wbTarget, SHEET_KANRI_MASTER) Then
        MsgBox "担当者マスタシート「" & SHEET_KANRI_MASTER & "」が見つかりません。", vbCritical, "シートエラー"
        GoTo FinalizeInit
    End If
    Set wsMaster = wbTarget.Sheets(SHEET_KANRI_MASTER)

    Me.担当者.List = wsMaster.Range("A2:A" & wsMaster.Cells(wsMaster.Rows.count, "A").End(xlUp).Row).Value

    Set m_CachedKoujiData = CreateObject("Scripting.Dictionary")

    Dim wsTarget As Worksheet
    Dim r As Long
    Dim currentStaff As String, currentKoujiName As String, currentKoujiBango As String

    If Not SheetExists(wbTarget, SHEET_KOUJI_LIST) Then
        MsgBox "データシート「" & SHEET_KOUJI_LIST & "」が見つかりません。", vbCritical, "シートエラー"
        GoTo FinalizeInit
    End If
    Set wsTarget = wbTarget.Sheets(SHEET_KOUJI_LIST)

    For r = wsTarget.Cells(wsTarget.Rows.count, "C").End(xlUp).Row To 2 Step -1
        currentStaff = Trim(CStr(wsTarget.Cells(r, "C").Value))
        currentKoujiBango = Trim(CStr(wsTarget.Cells(r, "D").Value))
        currentKoujiName = Trim(CStr(wsTarget.Cells(r, "E").Value))

        If currentStaff <> "" And currentKoujiName <> "" And currentKoujiBango <> "" Then
            If Not m_CachedKoujiData.Exists(currentStaff) Then
                m_CachedKoujiData.Add currentStaff, CreateObject("Scripting.Dictionary")
            End If
            If Not m_CachedKoujiData(currentStaff).Exists(currentKoujiBango) Then
                m_CachedKoujiData(currentStaff).Add currentKoujiBango, currentKoujiName
            Else
                m_CachedKoujiData(currentStaff).item(currentKoujiBango) = currentKoujiName
            End If
        End If
    Next r

FinalizeInit:
    If Not wsMaster Is Nothing Then Set wsMaster = Nothing
    If Not wsTarget Is Nothing Then Set wsTarget = Nothing
    If Not wbTarget Is Nothing Then wbTarget.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub

ErrorHandlerInit:
    MsgBox "初期化中にエラーが発生しました: " & Err.Description, vbCritical
    Resume FinalizeInit
End Sub

'--- キャッシュ変数 ---
Private m_CachedKoujiData As Object
Private m_isComboBoxUpdating As Boolean

'================================================================================
' 担当者変更時
'================================================================================
Private Sub 担当者_Change()
    If m_isComboBoxUpdating Then Exit Sub
    m_isComboBoxUpdating = True
    Me.工事名称.Clear
    Me.工事番号.Clear
    Me.削除.Enabled = False

    If Me.担当者.ListIndex = -1 Then
        m_isComboBoxUpdating = False
        Exit Sub
    End If

    Dim selectedStaff As String
    selectedStaff = Trim(Me.担当者.Value)

    If Not m_CachedKoujiData Is Nothing Then
        If m_CachedKoujiData.Exists(selectedStaff) Then
            Dim koujiDict As Object
            Set koujiDict = m_CachedKoujiData(selectedStaff)
            Dim koujiNamesArray() As String
            Dim koujiBangosArray() As String
            Dim count As Long
            count = 0
            Dim koujiBangoKey As Variant
            For Each koujiBangoKey In koujiDict.Keys
                count = count + 1
                ReDim Preserve koujiNamesArray(1 To count)
                ReDim Preserve koujiBangosArray(1 To count)
                koujiBangosArray(count) = koujiBangoKey
                koujiNamesArray(count) = koujiDict.item(koujiBangoKey)
            Next koujiBangoKey
            Me.工事名称.List = koujiNamesArray
            Me.工事番号.List = koujiBangosArray
        End If
    End If

    m_isComboBoxUpdating = False
End Sub

'================================================================================
' 工事名称変更時
'================================================================================
Private Sub 工事名称_Change()
    If m_isComboBoxUpdating Then Exit Sub
    m_isComboBoxUpdating = True
    Me.削除.Enabled = False
    Me.工事番号.Clear

    Dim selectedKoujiName As String
    selectedKoujiName = Trim(Me.工事名称.Value)
    Dim selectedStaff As String
    selectedStaff = Trim(Me.担当者.Value)

    If selectedStaff <> "" And Not m_CachedKoujiData Is Nothing Then
        If m_CachedKoujiData.Exists(selectedStaff) Then
            Dim koujiDict As Object
            Set koujiDict = m_CachedKoujiData(selectedStaff)
            Dim koujiBangoKey As Variant
            For Each koujiBangoKey In koujiDict.Keys
                If Trim(koujiDict.item(koujiBangoKey)) = selectedKoujiName Then
                    Me.工事番号.Value = koujiBangoKey
                    Me.削除.Enabled = True
                    Exit For
                End If
            Next koujiBangoKey
        End If
    End If

    m_isComboBoxUpdating = False
End Sub

'================================================================================
' 工事番号変更時
'================================================================================
Private Sub 工事番号_Change()
    If m_isComboBoxUpdating Then Exit Sub
    m_isComboBoxUpdating = True
    Me.削除.Enabled = False
    Me.工事名称.Clear

    Dim selectedKoujiBango As String
    selectedKoujiBango = Trim(Me.工事番号.Value)
    Dim selectedStaff As String
    selectedStaff = Trim(Me.担当者.Value)

    If selectedStaff <> "" And Not m_CachedKoujiData Is Nothing Then
        If m_CachedKoujiData.Exists(selectedStaff) Then
            Dim koujiDict As Object
            Set koujiDict = m_CachedKoujiData(selectedStaff)
            If koujiDict.Exists(selectedKoujiBango) Then
                Me.工事名称.Value = koujiDict.item(selectedKoujiBango)
                Me.削除.Enabled = True
            End If
        End If
    End If

    m_isComboBoxUpdating = False
End Sub

'================================================================================
' 「削除」ボタン
'================================================================================
Private Sub 削除_Click()
    Dim wbTarget As Workbook, wsTarget As Worksheet, wsMaster As Worksheet
    Dim rowToDelete As Long
    Dim r As Long
    Dim koujiNameToDelete As String, tantoushaToDelete As String, koujiBangoToDelete As String
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean

    If Me.担当者.ListIndex = -1 Then
        MsgBox "担当者が選択されていません。", vbExclamation
        Exit Sub
    End If
    If Trim(Me.工事名称.Value) = "" Or Trim(Me.工事番号.Value) = "" Then
        MsgBox "削除する工事名称または工事番号が選択されていません。", vbExclamation
        Exit Sub
    End If

    tantoushaToDelete = Me.担当者.Value
    koujiNameToDelete = Me.工事名称.Value
    koujiBangoToDelete = Me.工事番号.Value

    If MsgBox("担当者: " & tantoushaToDelete & vbCrLf & _
              "工事名称: " & koujiNameToDelete & vbCrLf & _
              "工事番号: " & koujiBangoToDelete & vbCrLf & vbCrLf & _
              "この工事情報を完全に削除します。よろしいですか？", _
              vbQuestion + vbYesNo + vbDefaultButton2, "削除の最終確認") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    originalEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo ErrorHandlerDelete

    Dim openedWbMain As Workbook
    For Each openedWbMain In Application.Workbooks
        If openedWbMain.FullName = m_TARGET_FILE_PATH Then
            MsgBox "対象のExcelファイルが既に開かれています。閉じてから再度実行してください。", vbCritical
            GoTo CleanUpDelete
        End If
    Next openedWbMain

    Set wbTarget = Application.Workbooks.Open(fileName:=m_TARGET_FILE_PATH, ReadOnly:=False, UpdateLinks:=0)

    If wbTarget Is Nothing Then
        MsgBox "対象ファイルを開けませんでした。他のユーザーが書き込みロックしている可能性があります。", vbCritical
        GoTo CleanUpDelete
    End If
    If wbTarget.ReadOnly Then
        MsgBox "対象ファイルは読み取り専用で開かれました。削除できません。他のユーザーが使用中か確認してください。", vbExclamation
        GoTo CleanUpDelete
    End If

    If Not SheetExists(wbTarget, SHEET_KOUJI_LIST) Then
        MsgBox "データシート「" & SHEET_KOUJI_LIST & "」が見つかりません。", vbCritical
        GoTo CleanUpDelete
    End If
    If Not SheetExists(wbTarget, SHEET_KANRI_MASTER) Then
        MsgBox "担当者マスタシート「" & SHEET_KANRI_MASTER & "」が見つかりません。", vbCritical
        GoTo CleanUpDelete
    End If

    Set wsTarget = wbTarget.Sheets(SHEET_KOUJI_LIST)
    Set wsMaster = wbTarget.Sheets(SHEET_KANRI_MASTER)

    rowToDelete = 0
    For r = wsTarget.Cells(wsTarget.Rows.count, "E").End(xlUp).Row To 2 Step -1
        If Trim(CStr(wsTarget.Cells(r, "C").Value)) = Trim(tantoushaToDelete) And _
           Trim(CStr(wsTarget.Cells(r, "E").Value)) = Trim(koujiNameToDelete) And _
           Trim(CStr(wsTarget.Cells(r, "D").Value)) = Trim(koujiBangoToDelete) Then
            rowToDelete = r
            Exit For
        End If
    Next r

    If rowToDelete = 0 Then
        MsgBox "削除対象のデータが見つかりませんでした。他のユーザーが既に削除した可能性があります。", vbExclamation
        GoTo CleanUpDelete
    End If

    Call SafeUnprotect(wsTarget)
    wsTarget.Rows(rowToDelete).Delete
    Call SafeProtectData(wsTarget)
    Call ClearAllFilters(wsTarget)
    wbTarget.Save
    Call UpdateLocalListSheet(wsTarget, wsMaster)
    MsgBox "削除が完了しました。", vbInformation

CleanUpDelete:
    If Not wsMaster Is Nothing Then Set wsMaster = Nothing
    If Not wsTarget Is Nothing Then Set wsTarget = Nothing
    If Not wbTarget Is Nothing Then wbTarget.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Unload Me
    Exit Sub

ErrorHandlerDelete:
    MsgBox "削除処理中にエラーが発生しました: " & Err.Description, vbCritical
    Resume CleanUpDelete
End Sub

'================================================================================
' ローカルシート更新
'================================================================================
Private Sub UpdateLocalListSheet(ByVal wsSource As Worksheet, ByVal wsMaster As Worksheet)
    Dim wsDest As Worksheet
    Dim destSheetName As String
    Dim lastRowSource As Long
    Dim copyRange As Range
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean

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

    On Error GoTo ErrorHandlerUpdateLocalDel
    Call SafeUnprotect(wsDest)
    wsDest.Range("A3:X" & wsDest.Rows.count).Clear
    lastRowSource = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    If lastRowSource >= 5 Then
        Set copyRange = wsSource.Range("A5:X" & lastRowSource)
        copyRange.Copy Destination:=wsDest.Range("A3")
    End If
    Call SafeProtectData(wsDest)
    Exit Sub

ErrorHandlerUpdateLocalDel:
    Call SafeProtectData(wsDest)
    MsgBox "ローカル一覧更新中にエラー: " & Err.Description, vbCritical

FinalizeUpdateLocal:
    Application.CutCopyMode = False
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Exit Sub

ErrorHandlerUpdateLocal:
    MsgBox "ローカルシート更新中に予期せぬエラー発生: " & Err.Description, vbCritical
    Resume FinalizeUpdateLocal
End Sub


