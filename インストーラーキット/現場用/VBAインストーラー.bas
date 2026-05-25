Attribute VB_Name = "VBAインストーラー"
Option Explicit

'================================================================================
' [ツール] VBAインストーラー（現場用）
'--------------------------------------------------------------------------------
' 【何のファイル？】
'   現場用xlsm に最新の VBA を一括取り込むツール。
'   フォーム(.frm)はコードのみ書き換え、コントロール（テキストボックス
'   やボタン配置）は壊さない v3 方式。
'
' 【手順】
'   1. xlsm を開く  2. Alt+F11  3. このファイルをインポート
'   4. InstallAll を実行  5. 現場用キットフォルダを選択
'================================================================================


'================================================================================
' VBA インストーラー（フォーム除外版）
'
' 【使い方】
' 1. このモジュールをVBAエディタにインポート
' 2. Alt+F8 でマクロ一覧を開き、「InstallAll」を実行
' 3. フォルダ選択ダイアログで作成キットのフォルダを選ぶ
' 4. 標準モジュール(.bas)とThisWorkbookだけが更新される
'
' 【重要】ユーザーフォーム(.frm)は上書きしない
' 　コントロール名のリネームやキャプション変更を保持するため。
' 　フォームを更新したい場合は手動でインポートしてください。
'
' 【フォルダ構成】
' 選んだフォルダに以下のファイルを置く：
'   設定用コード.bas
'   呼び出し.bas
'   PDF作成.bas
'   登録フォームを開く.bas
'   削除フォームを開く.bas
'   更新.bas
'   VBAインストーラー.bas （このモジュール自身）
'   ThisWorkbook.cls
'================================================================================

' インポート対象 .bas ファイル一覧（.frm は除外）
Private Const MODULES_LIST As String = _
    "設定用コード.bas|" & _
    "呼び出し.bas|" & _
    "PDF作成.bas|" & _
    "登録フォームを開く.bas|" & _
    "削除フォームを開く.bas|" & _
    "更新.bas"

'================================================================================
' メイン・ルーチン
'================================================================================
Public Sub InstallAll()
    Dim sourceFolder As String
    Dim files() As String
    Dim i As Long
    Dim successCount As Long, failCount As Long
    Dim logText As String

    ' --- VBEへのアクセス確認 ---
    If Not CheckVBEAccess() Then
        MsgBox "VBEへのプログラムからのアクセスが許可されていません。" & vbCrLf & vbCrLf & _
               "【ファイル】→【オプション】→【セキュリティセンター】→【セキュリティセンターの設定】→ " & vbCrLf & _
               "「マクロの設定」欄で、「VBAプロジェクトオブジェクトモデルへのアクセスを信頼」にチェックを入れてください。", _
               vbCritical, "インストールに失敗"
        Exit Sub
    End If

    ' --- フォルダ選択 ---
    sourceFolder = SelectFolder()
    If sourceFolder = "" Then Exit Sub
    If Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"

    ' --- 確認 ---
    If MsgBox("以下の処理を実行します：" & vbCrLf & vbCrLf & _
              "1. 既存の標準モジュール(.bas)を削除" & vbCrLf & _
              "2. 指定フォルダから.basファイルをインポート" & vbCrLf & _
              "3. ThisWorkbookのコードを更新" & vbCrLf & vbCrLf & _
              "※ユーザーフォーム(.frm)は上書きしません" & vbCrLf & vbCrLf & _
              "ソースフォルダ：" & vbCrLf & sourceFolder & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbQuestion, "インストール確認") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    files = Split(MODULES_LIST, "|")

    ' --- STEP 1: 既存の標準モジュールを削除（.basのみ） ---
    logText = "=== 削除処理 ===" & vbCrLf
    Dim removeFailCount As Long
    removeFailCount = 0
    For i = 0 To UBound(files)
        Dim moduleName As String
        Dim removeStatus As String
        moduleName = GetModuleName(files(i))
        ' 自分自身は削除しない（実行中に消えると止まる）
        If moduleName <> "VBAインストーラー" Then
            removeStatus = RemoveModuleWithStatus(moduleName)
            logText = logText & "  " & removeStatus & vbCrLf
            If Left(removeStatus, 1) = "×" Then removeFailCount = removeFailCount + 1
        End If
    Next i

    ' --- STEP 2: .basファイルをインポート ---
    logText = logText & vbCrLf & "=== インポート処理 ===" & vbCrLf
    For i = 0 To UBound(files)
        Dim filePath As String
        filePath = sourceFolder & files(i)
        Dim modName As String
        modName = GetModuleName(files(i))

        ' 自分自身はスキップ
        If modName = "VBAインストーラー" Then
            logText = logText & "  - スキップ（自身）: " & files(i) & vbCrLf
        ElseIf Dir(filePath) <> "" Then
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
            logText = logText & "  × ファイルが存在しません: " & files(i) & vbCrLf
        End If
    Next i

    ' --- STEP 3: ThisWorkbookの更新 ---
    logText = logText & vbCrLf & "=== ThisWorkbook更新 ===" & vbCrLf
    Dim twPath As String
    twPath = sourceFolder & "ThisWorkbook.cls"
    If Dir(twPath) <> "" Then
        If UpdateThisWorkbook(twPath) Then
            logText = logText & "  ○ ThisWorkbook 更新成功" & vbCrLf
        Else
            logText = logText & "  × ThisWorkbook 更新失敗" & vbCrLf
        End If
    Else
        logText = logText & "  - ThisWorkbook.cls がないのでスキップ" & vbCrLf
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' --- 処理結果表示 ---
    Dim resultIcon As Long
    Dim resultTitle As String
    If failCount = 0 And removeFailCount = 0 Then
        resultIcon = vbInformation
        resultTitle = "インストール完了"
    Else
        resultIcon = vbExclamation
        resultTitle = "インストール完了（エラーあり）"
    End If
    
    MsgBox "インストール処理が終了しました。" & vbCrLf & vbCrLf & _
           "削除失敗: " & removeFailCount & "モジュール" & vbCrLf & _
           "インポート成功: " & successCount & "モジュール" & vbCrLf & _
           "インポート失敗: " & failCount & "モジュール" & vbCrLf & vbCrLf & _
           "※フォームは変更していません" & vbCrLf & vbCrLf & _
           "詳細: " & vbCrLf & logText, _
           resultIcon, resultTitle
End Sub

'================================================================================
' フォルダ選択ダイアログ
'================================================================================
Private Function SelectFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "VBAソースフォルダを選択"
    fd.InitialFileName = ThisWorkbook.Path & "\"
    If fd.Show = -1 Then
        SelectFolder = fd.SelectedItems(1)
    Else
        SelectFolder = ""
    End If
End Function

'================================================================================
' VBEへのプログラムからのアクセスが有効かチェック
'================================================================================
Private Function CheckVBEAccess() As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject.VBComponents
    CheckVBEAccess = (Err.Number = 0)
    On Error GoTo 0
End Function

'================================================================================
' ファイル名からモジュール名を取得
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
' 指定名のモジュールを削除
'================================================================================
Private Function RemoveModule(ByVal moduleName As String) As Boolean
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Not comp Is Nothing Then
        ' フォーム（vbext_ct_MSForm）は誤って削除しないよう保護
        If comp.Type = 3 Then  ' 3 = vbext_ct_MSForm
            RemoveModule = False
            Exit Function
        End If
        ThisWorkbook.VBProject.VBComponents.Remove comp
        RemoveModule = (Err.Number = 0)
    Else
        RemoveModule = False
    End If
    Err.Clear
    On Error GoTo 0
End Function

' 削除のステータス（成功/失敗/スキップ）を文字列で返す
Private Function RemoveModuleWithStatus(ByVal moduleName As String) As String
    Dim comp As Object
    Dim errMsg As String
    
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Err.Number <> 0 Or comp Is Nothing Then
        RemoveModuleWithStatus = "- 存在せず: " & moduleName
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    
    ' フォーム保護
    If comp.Type = 3 Then  ' vbext_ct_MSForm
        RemoveModuleWithStatus = "- スキップ(フォーム): " & moduleName
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

'================================================================================
' ThisWorkbookコードを更新（.clsファイルからコードを取り出して直接挿入）
'================================================================================
Private Function UpdateThisWorkbook(ByVal clsPath As String) As Boolean
    Dim fileNum As Integer
    Dim fileContent As String
    Dim codeText As String

    fileNum = FreeFile
    Open clsPath For Input As #fileNum
    Do While Not EOF(fileNum)
        Dim line As String
        Line Input #fileNum, line
        fileContent = fileContent & line & vbCrLf
    Loop
    Close #fileNum

    ' ヘッダ属性行をスキップしてコード本体だけ取り出す
    Dim lines() As String
    lines = Split(fileContent, vbCrLf)
    Dim i As Long
    Dim inBody As Boolean
    inBody = False
    For i = 0 To UBound(lines)
        Dim trimmed As String
        trimmed = Trim(lines(i))
        If Not inBody Then
            If Left(trimmed, 9) = "VERSION 1" Or _
               Left(trimmed, 5) = "BEGIN" Or _
               Left(trimmed, 8) = "MultiUse" Or _
               trimmed = "END" Or _
               Left(trimmed, 9) = "Attribute" Then
                ' スキップ
            ElseIf trimmed = "" Then
                If codeText <> "" Then codeText = codeText & vbCrLf
            Else
                inBody = True
                codeText = codeText & lines(i) & vbCrLf
            End If
        Else
            codeText = codeText & lines(i) & vbCrLf
        End If
    Next i

    On Error Resume Next
    Dim tw As Object
    Set tw = ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
    tw.DeleteLines 1, tw.CountOfLines
    tw.AddFromString codeText
    UpdateThisWorkbook = (Err.Number = 0)
    On Error GoTo 0
End Function
