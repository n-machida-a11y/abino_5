Attribute VB_Name = "VBAインストーラー"
Option Explicit

'================================================================================
' VBA インストーラー（現場用・フォーム対応版 v2）
'
' 【使い方】
' 1. このモジュールをVBAエディタにインポート（初回のみ手動）
' 2. Alt+F8 でマクロ一覧を開き、「InstallAll」を実行
' 3. フォルダ選択ダイアログでキット（インストーラーキット\現場用）を選ぶ
' 4. 標準モジュール(.bas)・ユーザーフォーム(.frm)・ThisWorkbookが更新される
'
' 【v2の変更点】
' ・ユーザーフォーム(.frm)も一括置換するよう変更
'   （2026/5月リリースで 依頼書作成 等のフォームに変更があったため）
' ・確認ダイアログでフォームも置換することを明示
'
' 【フォルダに置くファイル】
'   設定用コード.bas
'   呼び出し.bas
'   PDF作成.bas
'   登録フォームを開く.bas
'   削除フォームを開く.bas
'   更新.bas
'   登録ボタン.frm (+ .frx)
'   削除フォーム.frm (+ .frx)
'   工事名称選択.frm (+ .frx)
'   再登録.frm (+ .frx)
'   依頼書作成.frm (+ .frx)
'   ThisWorkbook.cls
'   VBAインストーラー.bas  （このモジュール自身・自動スキップ）
'================================================================================

' 標準モジュール (.bas)
Private Const MODULES_LIST As String = _
    "設定用コード.bas|" & _
    "呼び出し.bas|" & _
    "PDF作成.bas|" & _
    "登録フォームを開く.bas|" & _
    "削除フォームを開く.bas|" & _
    "更新.bas"

' ユーザーフォーム (.frm) ※.frxは自動的に一緒にインポートされる
Private Const FORMS_LIST As String = _
    "登録ボタン.frm|" & _
    "削除フォーム.frm|" & _
    "工事名称選択.frm|" & _
    "再登録.frm|" & _
    "依頼書作成.frm"

'================================================================================
' メインルーチン
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
              "1. 既存の標準モジュール(.bas)を削除" & vbCrLf & _
              "2. 既存のユーザーフォーム(.frm)を削除" & vbCrLf & _
              "3. 指定フォルダから.basと.frmをインポート" & vbCrLf & _
              "4. ThisWorkbookのコードを更新" & vbCrLf & vbCrLf & _
              "★フォームも置換します★" & vbCrLf & _
              "（依頼書作成等のレイアウト変更を反映するため）" & vbCrLf & vbCrLf & _
              "ソースフォルダ:" & vbCrLf & sourceFolder & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "インストール確認") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    basFiles = Split(MODULES_LIST, "|")
    frmFiles = Split(FORMS_LIST, "|")

    ' --- STEP 1: 既存の標準モジュール削除 ---
    logText = "=== 標準モジュール削除 ===" & vbCrLf
    Dim basRemoveFail As Long
    For i = 0 To UBound(basFiles)
        Dim modName As String
        modName = GetModuleName(basFiles(i))
        If modName <> "VBAインストーラー" Then
            Dim s As String
            s = RemoveComponent(modName, False)
            logText = logText & "  " & s & vbCrLf
            If Left(s, 1) = "×" Then basRemoveFail = basRemoveFail + 1
        End If
    Next i

    ' --- STEP 2: 既存のフォーム削除 ---
    logText = logText & vbCrLf & "=== フォーム削除 ===" & vbCrLf
    Dim frmRemoveFail As Long
    For i = 0 To UBound(frmFiles)
        Dim frmName As String
        frmName = GetModuleName(frmFiles(i))
        Dim sf As String
        sf = RemoveComponent(frmName, True)
        logText = logText & "  " & sf & vbCrLf
        If Left(sf, 1) = "×" Then frmRemoveFail = frmRemoveFail + 1
    Next i

    ' --- STEP 3: .basインポート ---
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

    ' --- STEP 4: .frmインポート ---
    logText = logText & vbCrLf & "=== フォーム インポート ===" & vbCrLf
    For i = 0 To UBound(frmFiles)
        Dim frmPath As String
        frmPath = sourceFolder & frmFiles(i)
        Dim frxPath As String
        frxPath = sourceFolder & GetModuleName(frmFiles(i)) & ".frx"

        If Dir(frmPath) = "" Then
            frmFail = frmFail + 1
            logText = logText & "  × .frm無し: " & frmFiles(i) & vbCrLf
        ElseIf Dir(frxPath) = "" Then
            frmFail = frmFail + 1
            logText = logText & "  × .frx無し: " & GetModuleName(frmFiles(i)) & ".frx" & vbCrLf
        Else
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import frmPath
            If Err.Number = 0 Then
                frmSuccess = frmSuccess + 1
                logText = logText & "  ○ インポート: " & frmFiles(i) & vbCrLf
            Else
                frmFail = frmFail + 1
                logText = logText & "  × 失敗: " & frmFiles(i) & " (" & Err.Description & ")" & vbCrLf
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next i

    ' --- STEP 5: ThisWorkbook更新 ---
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
    If basFail = 0 And frmFail = 0 And basRemoveFail = 0 And frmRemoveFail = 0 Then
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
           "フォーム 削除失敗: " & frmRemoveFail & vbCrLf & _
           "フォーム インポート成功: " & frmSuccess & vbCrLf & _
           "フォーム インポート失敗: " & frmFail & vbCrLf & vbCrLf & _
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
' ファイル名からモジュール名を取得（拡張子除去）
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
' コンポーネント削除（標準モジュール/フォーム共通）
'   allowForm=True なら Type=3(フォーム)も削除する
'================================================================================
Private Function RemoveComponent(ByVal moduleName As String, ByVal allowForm As Boolean) As String
    Dim comp As Object
    Dim errMsg As String

    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Err.Number <> 0 Or comp Is Nothing Then
        RemoveComponent = "- 存在無: " & moduleName
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    ' フォームかどうか
    If comp.Type = 3 Then  ' vbext_ct_MSForm
        If Not allowForm Then
            RemoveComponent = "- フォーム保護スキップ: " & moduleName
            Exit Function
        End If
    End If

    Err.Clear
    ThisWorkbook.VBProject.VBComponents.Remove comp
    If Err.Number <> 0 Then
        errMsg = Err.Description
        RemoveComponent = "× 削除失敗: " & moduleName & " (" & errMsg & ")"
    Else
        RemoveComponent = "○ 削除: " & moduleName
    End If
    Err.Clear
    On Error GoTo 0
End Function

'================================================================================
' ThisWorkbookコードを更新（.clsファイルから本体抽出して挿入）
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
