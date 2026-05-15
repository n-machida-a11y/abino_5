Attribute VB_Name = "保存"
Option Explicit

Private Const CELL_INVOICE_FILE_NAME As String = "AK1"
Private Const CELL_ACCOUNTING_NO As String = "AF1"
Private Const CELL_MONTH_COL1 As String = "AE3"
Private Const CELL_MONTH_COL2 As String = "AB3"

'================================================================================
' 主要処理：保存＆印刷
'================================================================================
Public Sub 保存_印刷作業()
    Dim activeSht As Worksheet
    Set activeSht = ActiveSheet
    Dim initialScreen As Boolean
    initialScreen = Application.ScreenUpdating

    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    Call SyncInvoiceNoToMaster

    Dim baseFolder As String, invoiceFolder As String
    Dim fileName As String, monthYear As String

    fileName = CStr(activeSht.Range(CELL_INVOICE_FILE_NAME).Value)
    ' 年月フォルダ名を「YYYY年M月」形式で構築（旧: 月+年 のおかしな順序だった）
    Dim yearVal As String: yearVal = CStr(activeSht.Range(CELL_MONTH_COL2).Value)  ' AB3 = 年
    Dim monthVal As String: monthVal = CStr(activeSht.Range(CELL_MONTH_COL1).Value)  ' AE3 = 月
    If yearVal = "" Then yearVal = CStr(Year(Date))
    If monthVal = "" Then monthVal = CStr(Month(Date))
    monthYear = yearVal & "年" & monthVal & "月"

    ' ファイル名に使えない文字(\/:*?<>|)を除去
    fileName = SanitizeFileName(fileName)
    monthYear = SanitizeFileName(monthYear)

    If Trim(fileName) = "" Then
        MsgBox "ファイル名が取得できません。" & vbCrLf & _
               "請求書シートのAK1セルにファイル名が設定されているか確認してください。", vbCritical
        Exit Sub
    End If

    ' 保存先ベースフォルダ: 依頼検索!C1 セルの値を優先使用
    ' C1 が空の場合は Downloads にフォールバック
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

    ' 末尾が \ なら削除（連結時の二重 \ 防止）
    If Right(baseFolder, 1) = "\" Then baseFolder = Left(baseFolder, Len(baseFolder) - 1)

    invoiceFolder = baseFolder & "\" & monthYear & "請求書"

    ' FileSystemObjectで確実にフォルダ作成（OneDrive等でMkDirが失敗する環境対策）
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

    Dim newWb As Workbook
    Set newWb = Workbooks.Add
    activeSht.Copy Before:=newWb.Sheets(1)
    Application.DisplayAlerts = False  ' シート削除の確認ダイアログを抑制
    newWb.Sheets(1).Delete
    Application.DisplayAlerts = True

    Dim filePath As String
    filePath = invoiceFolder & "\" & fileName & ".xlsx"
    newWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    ' 保存後もシートは削除せず、そのまま残す
    ' (印刷し直しや再確認のため。不要になったらユーザー側で手動削除)

    Application.ScreenUpdating = initialScreen
    MsgBox "ファイルを保存しました。" & vbCrLf & "保存先：" & filePath, vbInformation
    Exit Sub

Cleanup:
    Application.ScreenUpdating = initialScreen
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
' マスタ同期処理：請求書番号をマスタ依頼履歴に書き込む
' 【修正】複数の早期 Exit を廃止し、Cleanup ラベルで一元管理する。
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

    ' 【変更】Application.Run("GetMasterPath") → GetMasterPath() の直接呼び出し
    mPath = GetMasterPath()

    On Error GoTo Cleanup

    ' 空パス対策：Dir("") はエラー52(ファイル名が不正)を投げる
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
    Set found = wsRireki.Columns("A").Find(What:=iraiNo, LookAt:=xlWhole)
    If found Is Nothing Then
        MsgBox "マスタ内に依頼NO [" & iraiNo & "] が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    targetRow = found.Row
    accountingNo = ActiveSheet.Range(CELL_ACCOUNTING_NO).Value
    wsRireki.Cells(targetRow, 2).Value = accountingNo

    wbMaster.Save

Cleanup:
    ' 正常時もエラー時も必ずここを通る
    If Not wbMaster Is Nothing Then wbMaster.Close False
    If Err.Number <> 0 Then
        MsgBox "マスタ同期中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    End If
End Sub

'================================================================================
' ファイル名に使えない文字を安全な記号で置換
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
