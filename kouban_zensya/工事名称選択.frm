VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 工事名称選択 
   Caption         =   "工事名称選択"
   ClientHeight    =   5676
   ClientLeft      =   72
   ClientTop       =   216
   ClientWidth     =   8484
   OleObjectBlob   =   "工事名称選択.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "工事名称選択"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ユーザーフォーム: 工事名称選択

Option Explicit

'================================================================================
' [フォーム] 工事名称選択（依頼書作成・再登録の前段の絞り込みフォーム）
'--------------------------------------------------------------------------------
' 【何のファイル？】
'   担当者を選択 → その担当者の工事名称リスト → さらに依頼書/再登録へ
'   進む2段絞り込みフォーム。
'
' 【しくみ】
'   - 担当者リストは担当者マスタから自部署フィルタ
'   - 工事一覧は工事番号の prefix（"03-" 等）で自部署のみ
'================================================================================


'================================================================================
' 【変更点】
'   - IS_TEST_MODE, m_TEST_FILE_PATH_FOR_TEST, m_PATH_CELL → 削除
'     （設定用コードモジュールの GetMasterPath() に一元化）
'   - m_KANRI_MASTER_NAME             → SHEET_KANRI_MASTER に置き換え
'   - MASTER_TARGET_SHEET_NAME_CELL   → CELL_TARGET_SHEET に置き換え
'   - ファイルパス取得ロジック         → GetMasterPath() に置き換え
'================================================================================

'--- このフォームを呼び出した元のプログラムに、選択結果を渡すための変数 ---
Public selectedKoujiName As String
Public SelectedTantousha As String
Public Cancelled As Boolean

'--- このフォーム内だけで使う変数 ---
Private m_CachedStaffList As Variant
Private m_CachedKoujiList As Object
Private m_selectedKoujiNameTemp As String

Private Sub Label3_Click()
End Sub

'================================================================================
' フォームが開かれる瞬間の準備処理
'================================================================================

Private Sub UserForm_Terminate()
    ' メモリリーク防止：Dictionary オブジェクトを明示的に解放
    Set m_CachedKoujiList = Nothing
End Sub

Private Sub UserForm_Initialize()
    Dim wbTarget_Init As Workbook
    Dim wsMaster_Init As Worksheet
    Dim targetSheetName As String

    Me.Cancelled = True
    Me.OK.Enabled = False
    Me.工事名称.Clear
    m_selectedKoujiNameTemp = ""
    Set m_CachedKoujiList = CreateObject("Scripting.Dictionary")

    On Error GoTo ErrorHandlerInit

    ' 【変更】個別のパス取得ロジック → GetMasterPath() に置き換え
    Dim targetFilePath As String
    targetFilePath = GetMasterPath()

    If Dir(targetFilePath) = "" Then
        MsgBox "指定ファイルが見つかりません。" & vbCrLf & targetFilePath, vbCritical
        Unload Me
        Exit Sub
    End If

    Set wbTarget_Init = Application.Workbooks.Open(fileName:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)

    ' 【2026/5/24 改修】担当者リストは担当者マスタから自部門フィルタで取得
    m_CachedStaffList = GetStaffListFiltered(wbTarget_Init, GetMyDeptCode())
    If IsEmpty(m_CachedStaffList) Then
        MsgBox "担当者マスタに自部門(" & GetMyDeptCode() & ")の担当者が見つかりません。", vbExclamation
        GoTo FinalizeInit
    End If
    Me.担当者.List = m_CachedStaffList

    ' 工事番号一覧シート名を管理マスタA列「工事番号一覧シート」項目から取得
    targetSheetName = GetSheetNameFromMaster(wbTarget_Init, "工事番号一覧シート", SHEET_KOUJI_LIST)
    If targetSheetName = "" Then
        MsgBox "「管理マスタ」シートの「工事番号一覧シート」項目に値が設定されていません。", vbCritical
        GoTo FinalizeInit
    End If

    Dim wsTarget As Worksheet
    If Not SheetExists(wbTarget_Init, targetSheetName) Then
        MsgBox "外部ファイルにシート「" & targetSheetName & "」が見つかりません。", vbCritical
        GoTo FinalizeInit
    End If
    Set wsTarget = wbTarget_Init.Sheets(targetSheetName)

    Dim r As Long
    Dim currentStaff As String, currentKoujiName As String, currentKoujiBango As String

    ' 自部門コードでフィルタするための prefix
    Dim deptPrefix As String
    On Error Resume Next
    deptPrefix = GetMyDeptCode() & "-"
    On Error GoTo 0
    If deptPrefix = "-" Then deptPrefix = "03-"

    For r = wsTarget.Cells(wsTarget.Rows.count, "C").End(xlUp).Row To 2 Step -1
        currentStaff = Trim(CStr(wsTarget.Cells(r, "C").Value))
        currentKoujiBango = Trim(CStr(wsTarget.Cells(r, "D").Value))
        currentKoujiName = Trim(CStr(wsTarget.Cells(r, "E").Value))
        ' 部門フィルタ: 自部門の工事番号のみ
        If Left(currentKoujiBango, Len(deptPrefix)) <> deptPrefix Then GoTo NextRow_Sel
        If currentStaff <> "" And currentKoujiName <> "" Then
            If Not m_CachedKoujiList.Exists(currentStaff) Then
                m_CachedKoujiList.Add currentStaff, CreateObject("Scripting.Dictionary")
            End If
            If Not m_CachedKoujiList(currentStaff).Exists(currentKoujiName) Then
                m_CachedKoujiList(currentStaff).Add currentKoujiName, True
            End If
        End If
NextRow_Sel:
    Next r

FinalizeInit:
    If Not wbTarget_Init Is Nothing Then wbTarget_Init.Close SaveChanges:=False
    Exit Sub

ErrorHandlerInit:
    MsgBox "初期化中に予期せぬエラー発生: " & Err.Description, vbCritical
    Resume FinalizeInit
End Sub

'================================================================================
' 「担当者」ドロップダウンが変更されたときの処理
'================================================================================
Private Sub 担当者_Change()
    Me.工事名称.Clear
    Me.OK.Enabled = False
    m_selectedKoujiNameTemp = ""

    If Me.担当者.ListIndex = -1 Then Exit Sub

    Dim selectedStaff As String
    selectedStaff = Trim(Me.担当者.Value)

    If Not m_CachedKoujiList Is Nothing Then
        If m_CachedKoujiList.Exists(selectedStaff) Then
            Me.工事名称.List = m_CachedKoujiList(selectedStaff).Keys
        End If
    End If
End Sub

'================================================================================
' 「工事名称」リストボックスがクリックされたときの処理
'================================================================================
Private Sub 工事名称_Click()
    If Me.工事名称.ListIndex > -1 Then
        m_selectedKoujiNameTemp = Me.工事名称.Value
        Me.OK.Enabled = True
    Else
        m_selectedKoujiNameTemp = ""
        Me.OK.Enabled = False
    End If
End Sub

'================================================================================
' 「OK」ボタンがクリックされたときの処理
'================================================================================
Private Sub OK_Click()
    If m_selectedKoujiNameTemp = "" Then
        MsgBox "工事名称が選択されていません。", vbExclamation
        Exit Sub
    End If

    Me.selectedKoujiName = m_selectedKoujiNameTemp
    Me.SelectedTantousha = Me.担当者.Value
    Me.Cancelled = False
    Me.Hide
End Sub

'================================================================================
' 「×」ボタンで閉じられたときの処理
'================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Cancelled = True
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
