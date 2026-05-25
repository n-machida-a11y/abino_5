Attribute VB_Name = "VBAインストーラー"
Option Explicit

'================================================================================
' VBA インストーラー（経理用）
'
' 【対象】【共有_経理用】請求書作成.xlsm
'
' 【使い方】
' 1. このモジュールをVBAエディタにインポート（初回のみ手動）
' 2. Alt+F8 でマクロ一覧を開き、「InstallAll」を実行
' 3. フォルダ選択ダイアログでキット（インストーラーキット\経理用）を選ぶ
' 4. 標準モジュール(.bas)が更新される
'
' 【インストールするモジュール】
'   設定用コード.bas   (VB_Name="Config")
'   シート更新.bas
'   保存.bas
'   請求書作成.bas
'
' 【注意】
' ・ユーザーフォーム(.frm)は経理用にはありません
' ・依頼検索シートのG1セル(マスタパス)・C1セル(保存先パス)の値は変更しません
'================================================================================

Private Const MODULES_LIST As String = _
    "設定用コード.bas|" & _
    "シート更新.bas|" & _
    "保存.bas|" & _
    "請求書作成.bas"

' VB_Nameが特殊なファイルのマッピング（ファイル名 → 実モジュール名）
' 設定用コード.bas は内部VB_Nameが "Config" になっているので削除時はこちら
Private Const MODULE_NAME_OVERRIDES As String = _
    "設定用コード.bas=Config"

'================================================================================
' メインルーチン
'================================================================================
Public Sub InstallAll()
    Dim sourceFolder As String
    Dim files() As String
    Dim i As Long
    Dim successCount As Long, failCount As Long, removeFailCount As Long
    Dim logText As String

    If Not CheckVBEAccess() Then
        MsgBox "VBEへのプログラムからのアクセスが許可されていません。" & vbCrLf & vbCrLf & _
               "【ファイル】→【オプション】→【セキュリティセンター】→【セキュリティセンターの設定】→ " & vbCrLf & _
               "「マクロの設定」欄で、「VBAプロジェクトオブジェクトモデルへのアクセスを信頼」にチェックを入れてください。", _
               vbCritical, "インストールに失敗"
        Exit Sub
    End If

    sourceFolder = SelectFolder()
    If sourceFolder = "" Then Exit Sub
    If Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"

    If MsgBox("以下の処理を実行します:" & vbCrLf & vbCrLf & _
              "1. 既存の標準モジュール(.bas)を削除" & vbCrLf & _
              "2. 指定フォルダから.basファイルをインポート" & vbCrLf & vbCrLf & _
              "★依頼検索シートのG1/C1の値は変更しません★" & vbCrLf & vbCrLf & _
              "ソースフォルダ:" & vbCrLf & sourceFolder & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "インストール確認") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    files = Split(MODULES_LIST, "|")

    ' --- STEP 1: 既存モジュール削除 ---
    logText = "=== 削除処理 ===" & vbCrLf
    For i = 0 To UBound(files)
        Dim modName As String
        modName = ResolveModuleName(files(i))
        Dim s As String
        s = RemoveModuleWithStatus(modName)
        logText = logText & "  " & s & vbCrLf
        If Left(s, 1) = "×" Then removeFailCount = removeFailCount + 1
    Next i

    ' --- STEP 2: インポート ---
    logText = logText & vbCrLf & "=== インポート処理 ===" & vbCrLf
    For i = 0 To UBound(files)
        Dim filePath As String
        filePath = sourceFolder & files(i)
        If Dir(filePath) <> "" Then
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import filePath
            If Err.Number = 0 Then
                successCount = successCount + 1
                logText = logText & "  ○ インポート: " & files(i) & vbCrLf
            Else
                failCount = failCount + 1
                logText = logText & "  × 失敗: " & files(i) & " (" & Err.Description & ")" & vbCrLf
            End If
            Err.Clear
            On Error GoTo 0
        Else
            failCount = failCount + 1
            logText = logText & "  × ファイル無し: " & files(i) & vbCrLf
        End If
    Next i

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim resultIcon As Long, resultTitle As String
    If failCount = 0 And removeFailCount = 0 Then
        resultIcon = vbInformation
        resultTitle = "インストール完了"
    Else
        resultIcon = vbExclamation
        resultTitle = "インストール完了（エラーあり）"
    End If

    MsgBox "インストール処理を終了しました。" & vbCrLf & vbCrLf & _
           "削除失敗: " & removeFailCount & "モジュール" & vbCrLf & _
           "インポート成功: " & successCount & "モジュール" & vbCrLf & _
           "インポート失敗: " & failCount & "モジュール" & vbCrLf & vbCrLf & _
           "詳細:" & vbCrLf & logText, _
           resultIcon, resultTitle
End Sub

'================================================================================
' ファイル名から実モジュール名を解決
'   オーバーライドマップにあればそれを使い、なければ拡張子除去
'================================================================================
Private Function ResolveModuleName(ByVal fileName As String) As String
    Dim overrides() As String, pair() As String, i As Long
    overrides = Split(MODULE_NAME_OVERRIDES, "|")
    For i = 0 To UBound(overrides)
        If Len(overrides(i)) > 0 Then
            pair = Split(overrides(i), "=")
            If UBound(pair) = 1 Then
                If pair(0) = fileName Then
                    ResolveModuleName = pair(1)
                    Exit Function
                End If
            End If
        End If
    Next i
    ResolveModuleName = GetModuleName(fileName)
End Function

'================================================================================
' フォルダ選択ダイアログ
'================================================================================
Private Function SelectFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "VBAソースフォルダを選ぶ"
    fd.InitialFileName = ThisWorkbook.Path & "\"
    If fd.Show = -1 Then
        SelectFolder = fd.SelectedItems(1)
    Else
        SelectFolder = ""
    End If
End Function

'================================================================================
' VBEへのアクセス確認
'================================================================================
Private Function CheckVBEAccess() As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject.VBComponents
    CheckVBEAccess = (Err.Number = 0)
    On Error GoTo 0
End Function

'================================================================================
' ファイル名からモジュール名（拡張子除去）
'================================================================================
Private Function GetModuleName(ByVal fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetModuleName = Left(fileName, dotPos - 1)
    Else
        GetModuleName = fileName
    End If
End Function

'================================================================================
' 指定名のモジュール削除（ステータス文字列を返す）
'================================================================================
Private Function RemoveModuleWithStatus(ByVal moduleName As String) As String
    Dim comp As Object
    Dim errMsg As String

    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Err.Number <> 0 Or comp Is Nothing Then
        RemoveModuleWithStatus = "- 存在無: " & moduleName
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If comp.Type = 3 Then  ' vbext_ct_MSForm
        RemoveModuleWithStatus = "- フォーム保護スキップ: " & moduleName
        Exit Function
    End If

    Err.Clear
    ThisWorkbook.VBProject.VBComponents.Remove comp
    If Err.Number <> 0 Then
        errMsg = Err.Description
        RemoveModuleWithStatus = "× 削除失敗: " & moduleName & " (" & errMsg & ")"
    Else
        RemoveModuleWithStatus = "○ 削除: " & moduleName
    End If
    Err.Clear
    On Error GoTo 0
End Function
