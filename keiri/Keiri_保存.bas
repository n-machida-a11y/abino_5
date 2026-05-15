Option Explicit

'================================================================================
' 設定エリア
'================================================================================
' 請求書シート側の管理セル
Private Const CELL_FILENAME As String = "AK1"
Private Const CELL_KEIRI_NO As String = "AF1"
Private Const CELL_MONTH As String = "AE3"
' 年のセル（請求書シートの年表示セル）
Private Const CELL_YEAR As String = "AB3"

' 依頼検索シート側の保存先パス記載セル (C1)
Private Const SRC_FOLDER_PATH_CELL As String = "C1"

' ※予備のマスターファイルパス設定（BACKUP_MASTER_PATH）は削除しました。
'   マスターパスは必ずメインモジュール経由またはシート設定から取得します。

'================================================================================
' 保存処理（クラッシュ防止・安定版）：Excelのみ保存
'================================================================================
Sub 保存_印刷作業()
    Dim targetSheet As Worksheet
    Dim baseFolder As String
    Dim monthFolder As String
    Dim fileName As String
    Dim fullPathExcel As String
    Dim keiriNo As String
    Dim fso As Object
    Dim newWb As Workbook
    
    ' 1. オブジェクトのセット
    Set targetSheet = ActiveSheet
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 安全装置：システムシートは除外
    Select Case targetSheet.Name
        Case "依頼検索", "作成選択", "入力画面", "依頼履歴", "部署マスタ", "テンプレート"
            MsgBox "システムシートは保存できません。作成された請求書シートを表示して実行してください。", vbExclamation
            Exit Sub
    End Select

    ' 2. 保存情報の取得
    On Error Resume Next
    baseFolder = Trim(ThisWorkbook.Worksheets("依頼検索").Range(SRC_FOLDER_PATH_CELL).Value)
    
    fileName = Trim(targetSheet.Range(CELL_FILENAME).Value)
    keiriNo = Trim(targetSheet.Range(CELL_KEIRI_NO).Value)
    
    ' 年と月を取得して「〇〇年〇月請求書」形式にする
    Dim yVal As String: yVal = targetSheet.Range(CELL_YEAR).Value
    Dim mVal As String: mVal = targetSheet.Range(CELL_MONTH).Value
    
    ' もし空欄なら現在の日付から補完
    If yVal = "" Then yVal = Year(Date)
    If mVal = "" Then mVal = Month(Date)
    
    monthFolder = yVal & "年" & mVal & "月請求書"
    On Error GoTo 0
    
    If baseFolder = "" Or fileName = "" Then
        MsgBox "保存設定が不完全です。" & vbCrLf & _
               "・依頼検索シートの" & SRC_FOLDER_PATH_CELL & "セル（保存先パス）" & vbCrLf & _
               "・請求書シートのAK1セル（ファイル名）" & vbCrLf & _
               "を確認してください。", vbExclamation
        Exit Sub
    End If
    
    ' 3. パス生成とフォルダ作成
    If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
    Dim targetFolder As String: targetFolder = baseFolder & monthFolder & "\"
    
    On Error Resume Next
    If Not fso.FolderExists(baseFolder) Then
        MsgBox "基本フォルダが見つかりません：" & vbCrLf & baseFolder & vbCrLf & "フォルダが存在するか確認してください。", vbCritical
        Exit Sub
    End If
    If Not fso.FolderExists(targetFolder) Then
        fso.CreateFolder targetFolder
    End If
    On Error GoTo 0
    
    fullPathExcel = targetFolder & fileName & ".xlsx"

    ' 4. 実行確認
    If MsgBox("月別フォルダ「" & monthFolder & "」へ発行しますか？" & vbCrLf & _
              "※Excelファイルのみ保存します。", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    ' 5. マスター側への同期（詳細なエラーチェック付き）
    If SyncInvoiceNoToMaster(targetSheet) = False Then
        If MsgBox("マスターへの同期に失敗しましたが、保存を続行しますか？" & vbCrLf & _
                  "（「いいえ」を押すと中断し、原因を確認できます）", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    ' 6. 出力処理
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo SaveError
    
    targetSheet.Copy
    Set newWb = ActiveWorkbook
    
    ' 既存ファイルがあっても上書き保存
    newWb.SaveAs fileName:=fullPathExcel, FileFormat:=xlOpenXMLWorkbook
    
    newWb.Close SaveChanges:=False
    
    ' 7. 元のシートを削除
    targetSheet.Delete
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "発行が完了しました。" & vbCrLf & "保存先：" & fullPathExcel, vbInformation
    Exit Sub

SaveError:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "エラーにより中断しました。以下を確認してください：" & vbCrLf & _
           "1. 保存先のExcelファイルが開いたままではないか" & vbCrLf & _
           "2. 保存先パスに禁止文字が含まれていないか" & vbCrLf & _
           "エラー内容：" & Err.Description, vbCritical
End Sub

' 外部マスター同期
Private Function SyncInvoiceNoToMaster(ByVal wsInv As Worksheet) As Boolean
    Dim mPath As String: Dim wbM As Workbook: Dim wsM As Worksheet: Dim irNo As String: Dim f As Range
    SyncInvoiceNoToMaster = False
    
    ' 依頼検索シートのA2セル（依頼No）を取得
    irNo = Trim(ThisWorkbook.Sheets("依頼検索").Range("A2").Value)
    
    ' パスを取得（取得失敗時はメッセージを出して終了）
    mPath = GetMasterPath()
    
    If mPath = "" Then
        MsgBox "マスターファイルのパス設定が空です。" & vbCrLf & "Mainモジュールまたはコード内の設定を確認してください。", vbExclamation
        Exit Function
    End If
    
    If Dir(mPath) = "" Then
        MsgBox "マスターファイルが見つかりません。" & vbCrLf & _
               "以下のパスを確認してください：" & vbCrLf & mPath, vbExclamation
        Exit Function
    End If
    
    On Error Resume Next
    Set wbM = Workbooks(Dir(mPath))
    If wbM Is Nothing Then
        Set wbM = Workbooks.Open(fileName:=mPath, ReadOnly:=False, UpdateLinks:=0)
    End If
    On Error GoTo 0
    
    If wbM Is Nothing Then
        MsgBox "マスターファイルを開けませんでした。" & vbCrLf & mPath, vbExclamation
        Exit Function
    End If
    
    Set wsM = wbM.Sheets("依頼履歴")
    Set f = wsM.Columns("A").Find(What:=irNo, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not f Is Nothing Then
        ' B列に経理番号を書き戻し
        wsM.Cells(f.Row, "B").Value = wsInv.Range(CELL_KEIRI_NO).Value
        wbM.Save
        SyncInvoiceNoToMaster = True
    Else
        MsgBox "マスターファイル内に、この依頼NO(" & irNo & ")が見つかりませんでした。" & vbCrLf & _
               "先に「新規作成」等でマスターに登録されているか確認してください。", vbExclamation
    End If
    
    wbM.Close SaveChanges:=False
End Function

Private Function GetMasterPath() As String
    Dim path As String
    ' ①Main_Search_Update.vbaにあるGetMasterPath関数を呼び出し試行
    On Error Resume Next
    path = Application.Run("GetMasterPath")
    On Error GoTo 0
    
    ' ※以前あった予備パス(BACKUP_MASTER_PATH)へのフォールバック処理は削除しました。
    '   取得できなかった場合(path="")は、そのまま空文字を返し、呼び出し元でエラーとします。
    
    GetMasterPath = path
End Function

Sub 印刷()
    On Error Resume Next
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub

