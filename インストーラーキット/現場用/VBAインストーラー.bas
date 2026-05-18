Attribute VB_Name = "VBAインストーラー"
Option Explicit

'================================================================================
' VBA インストーラー（現場用・フォーム保持版 v3）
'
' 【v3の変更点】
' ・v2 までは .frm を完全置換していたが、フォームのコントロール定義
'   （.frx バイナリ）まで置き換わってしまい、コントロール欠落の事故が
'   発生したため、v3 では「フォームは消さずに、コード本体だけ書き換え」
'   方式に変更。
' ・お客様の手元で既に正しく動作しているフォームのコントロール定義は
'   無傷で保持される。
'
' 【処理内容】
' 1. 標準モジュール(.bas)は従来どおり削除＆インポート
' 2. ユーザーフォーム(.frm)は削除せず、そのフォームのコード本体だけを
'    .frm ファイルから抽出して、既存フォームのCodeModuleに書き込み
' 3. ThisWorkbook も従来どおりコード本体だけ書き換え
'
' 【使い方】
' 1. このモジュールをVBAエディタにインポート
' 2. Alt+F8 で「InstallAll」を実行
' 3. フォルダ選択ダイアログでキットフォルダを選ぶ
'================================================================================

' 標準モジュール (.bas) ― 完全置換
Private Const MODULES_LIST As String = _
    "設定用コード.bas|" & _
    "呼び出し.bas|" & _
    "PDF作成.bas|" & _
    "登録フォームを開く.bas|" & _
    "削除フォームを開く.bas|" & _
    "更新.bas"

' ユーザーフォーム ― コード本体のみ書き換え（.frx は触らない）
Private Const FORMS_LIST As String = _
    "登録ボタン.frm|" & _
    "削除フォーム.frm|" & _
    "工事名称選択.frm|" & _
    "再登録.frm|" & _
    "依頼書作成.frm"

'================================================================================
Public Sub InstallAll()
    Dim sourceFolder As String
    Dim basFiles() As String, frmFiles() As String
    Dim i As Long
    Dim basSuccess As Long, basFail As Long
    Dim frmSuccess As Long, frmFail As Long
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
              "1. 既存の標準モジュール(.bas)を削除し、新しいモジュールをインポート" & vbCrLf & _
              "2. ユーザーフォームのコード本体だけ最新版に書き換え" & vbCrLf & _
              "   （フォームのレイアウト・コントロールは無傷で保持）" & vbCrLf & _
              "3. ThisWorkbookのコードを更新" & vbCrLf & vbCrLf & _
              "ソースフォルダ:" & vbCrLf & sourceFolder & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "インストール確認") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    basFiles = Split(MODULES_LIST, "|")
    frmFiles = Split(FORMS_LIST, "|")

    ' --- STEP 1: 既存の標準モジュールを削除 ---
    logText = "=== 標準モジュール削除 ===" & vbCrLf
    Dim basRemoveFail As Long
    For i = 0 To UBound(basFiles)
        Dim modName As String
        modName = GetModuleName(basFiles(i))
        If modName <> "VBAインストーラー" Then
            Dim s As String
            s = RemoveStandardModule(modName)
            logText = logText & "  " & s & vbCrLf
            If Left(s, 1) = "×" Then basRemoveFail = basRemoveFail + 1
        End If
    Next i

    ' --- STEP 2: 標準モジュールをインポート ---
    logText = logText & vbCrLf & "=== 標準モジュール インポート ===" & vbCrLf
    For i = 0 To UBound(basFiles)
        Dim filePath As String
        filePath = sourceFolder & basFiles(i)
        Dim modName2 As String
        modName2 = GetModuleName(basFiles(i))

        If modName2 = "VBAインストーラー" Then
            logText = logText & "  - スキップ(自身): " & basFiles(i) & vbCrLf
        ElseIf Dir(filePath) <> "" Then
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import filePath
            If Err.Number = 0 Then
                basSuccess = basSuccess + 1
                logText = logText & "  ○ インポート: " & basFiles(i) & vbCrLf
            Else
                basFail = basFail + 1
                logText = logText & "  × 失敗: " & basFiles(i) & " (" & Err.Description & ")" & vbCrLf
            End If
            Err.Clear
            On Error GoTo 0
        Else
            basFail = basFail + 1
            logText = logText & "  × ファイル無し: " & basFiles(i) & vbCrLf
        End If
    Next i

    ' --- STEP 3: フォームのコード本体だけ書き換え（.frx/レイアウトは触らない） ---
    logText = logText & vbCrLf & "=== フォーム コード書き換え ===" & vbCrLf
    For i = 0 To UBound(frmFiles)
        Dim frmPath As String
        frmPath = sourceFolder & frmFiles(i)
        Dim frmName As String
        frmName = GetModuleName(frmFiles(i))

        If Dir(frmPath) = "" Then
            frmFail = frmFail + 1
            logText = logText & "  × .frm無し: " & frmFiles(i) & vbCrLf
        Else
            Dim r As String
            r = UpdateFormCode(frmName, frmPath)
            logText = logText & "  " & r & vbCrLf
            If Left(r, 1) = "×" Then
                frmFail = frmFail + 1
            Else
                frmSuccess = frmSuccess + 1
            End If
        End If
    Next i

    ' --- STEP 4: ThisWorkbook更新 ---
    logText = logText & vbCrLf & "=== ThisWorkbook更新 ===" & vbCrLf
    Dim twPath As String
    twPath = sourceFolder & "ThisWorkbook.cls"
    If Dir(twPath) <> "" Then
        If UpdateThisWorkbook(twPath) Then
            logText = logText & "  ○ ThisWorkbook 更新完了" & vbCrLf
        Else
            logText = logText & "  × ThisWorkbook 更新失敗" & vbCrLf
        End If
    Else
        logText = logText & "  - ThisWorkbook.cls 無し（スキップ）" & vbCrLf
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim resultIcon As Long, resultTitle As String
    If basFail = 0 And frmFail = 0 And basRemoveFail = 0 Then
        resultIcon = vbInformation
        resultTitle = "インストール完了"
    Else
        resultIcon = vbExclamation
        resultTitle = "インストール完了（エラーあり）"
    End If

    MsgBox "インストール処理を終了しました。" & vbCrLf & vbCrLf & _
           "標準モジュール 削除失敗: " & basRemoveFail & vbCrLf & _
           "標準モジュール インポート成功: " & basSuccess & vbCrLf & _
           "標準モジュール インポート失敗: " & basFail & vbCrLf & _
           "フォームコード 書き換え成功: " & frmSuccess & vbCrLf & _
           "フォームコード 書き換え失敗: " & frmFail & vbCrLf & vbCrLf & _
           "詳細:" & vbCrLf & logText, _
           resultIcon, resultTitle
End Sub

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
' ファイル名 → モジュール名（拡張子除去）
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
' 標準モジュール削除（フォームは保護してスキップ）
'================================================================================
Private Function RemoveStandardModule(ByVal moduleName As String) As String
    Dim comp As Object
    Dim errMsg As String

    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Err.Number <> 0 Or comp Is Nothing Then
        RemoveStandardModule = "- 存在無: " & moduleName
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    ' フォームは絶対に削除しない（v3 のキモ）
    If comp.Type = 3 Then  ' vbext_ct_MSForm
        RemoveStandardModule = "- フォーム保護スキップ: " & moduleName
        Exit Function
    End If

    Err.Clear
    ThisWorkbook.VBProject.VBComponents.Remove comp
    If Err.Number <> 0 Then
        errMsg = Err.Description
        RemoveStandardModule = "× 削除失敗: " & moduleName & " (" & errMsg & ")"
    Else
        RemoveStandardModule = "○ 削除: " & moduleName
    End If
    Err.Clear
    On Error GoTo 0
End Function

'================================================================================
' フォームのコード本体だけ書き換え（.frx・レイアウトは触らない）
'   .frmファイルから「Attribute以降」のコード本体を抽出して、
'   既存フォームのCodeModuleに書き込む。
'================================================================================
Private Function UpdateFormCode(ByVal formName As String, ByVal frmPath As String) As String
    ' 既存フォームの存在確認
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(formName)
    On Error GoTo 0
    If comp Is Nothing Then
        UpdateFormCode = "× フォーム存在せず: " & formName
        Exit Function
    End If
    If comp.Type <> 3 Then  ' MSForm
        UpdateFormCode = "× フォームではない: " & formName
        Exit Function
    End If

    ' .frm ファイルからコード本体を抽出
    Dim codeText As String
    codeText = ExtractCodeBody(frmPath)
    If codeText = "" Then
        UpdateFormCode = "× コード抽出失敗: " & formName
        Exit Function
    End If

    ' CodeModule に書き込み
    On Error Resume Next
    Dim cm As Object
    Set cm = comp.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText
    If Err.Number <> 0 Then
        UpdateFormCode = "× コード書き込み失敗: " & formName & " (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    UpdateFormCode = "○ コード書き換え: " & formName
End Function

'================================================================================
' .frm/.cls ファイルからコード本体だけ抽出（ヘッダー部はスキップ）
'   VERSION 5.00 / Begin{}End / Attribute VB_xxx 行はスキップして、
'   実コード（Option Explicit以降）から最後までを返す。
'================================================================================
Private Function ExtractCodeBody(ByVal filePath As String) As String
    Dim fileNum As Integer
    Dim fileContent As String

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Dim line As String
        Line Input #fileNum, line
        fileContent = fileContent & line & vbCrLf
    Loop
    Close #fileNum

    Dim lines() As String
    lines = Split(fileContent, vbCrLf)
    Dim i As Long
    Dim inBody As Boolean
    Dim inBeginBlock As Boolean
    Dim beginDepth As Long
    Dim codeText As String
    inBody = False
    inBeginBlock = False
    beginDepth = 0

    For i = 0 To UBound(lines)
        Dim trimmed As String
        trimmed = Trim(lines(i))

        If Not inBody Then
            ' ヘッダー部のスキップ判定
            If Left(trimmed, 9) = "VERSION 1" Or Left(trimmed, 9) = "VERSION 5" Then
                ' バージョン宣言：スキップ
            ElseIf Left(trimmed, 6) = "Begin " Or Left(trimmed, 5) = "BEGIN" Then
                ' Begin ブロック開始：スキップ開始
                inBeginBlock = True
                beginDepth = beginDepth + 1
            ElseIf inBeginBlock And Left(trimmed, 3) = "End" Then
                ' End で Begin ブロック終了
                beginDepth = beginDepth - 1
                If beginDepth <= 0 Then
                    inBeginBlock = False
                    beginDepth = 0
                End If
            ElseIf inBeginBlock Then
                ' Begin ブロック内：スキップ
            ElseIf Left(trimmed, 9) = "Attribute" Then
                ' Attribute 行：スキップ
            ElseIf trimmed = "" Then
                ' 空行：本文に入る前はスキップ
            Else
                ' これ以降は本文
                inBody = True
                codeText = codeText & lines(i) & vbCrLf
            End If
        Else
            codeText = codeText & lines(i) & vbCrLf
        End If
    Next i

    ExtractCodeBody = codeText
End Function

'================================================================================
' ThisWorkbookコードを更新
'================================================================================
Private Function UpdateThisWorkbook(ByVal clsPath As String) As Boolean
    Dim codeText As String
    codeText = ExtractCodeBody(clsPath)
    If codeText = "" Then
        UpdateThisWorkbook = False
        Exit Function
    End If

    On Error Resume Next
    Dim tw As Object
    Set tw = ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
    tw.DeleteLines 1, tw.CountOfLines
    tw.AddFromString codeText
    UpdateThisWorkbook = (Err.Number = 0)
    On Error GoTo 0
End Function
