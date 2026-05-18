Attribute VB_Name = "VBAインストーラー"
Option Explicit

'================================================================================
' VBA インストーラー（管理者用）
'
' 【使い方】
' 1. このモジュールをVBAエディタにインポート（初回のみ手動）
' 2. Alt+F8 でマクロ一覧を開き、「InstallAll」を実行
' 3. フォルダ選択ダイアログで管理者用コード一式の入ったフォルダを選ぶ
' 4. 標準モジュール(.bas)が更新される
'
' 【フォルダに置くファイル】
'   管理者設定用コード.bas
'   管理者操作.bas
'   VBAインストーラー.bas  （このモジュール自身・自動スキップ）
'================================================================================

Private Const MODULES_LIST As String = _
    "管理者設定用コード.bas|" & _
    "管理者操作.bas"

Public Sub InstallAll()
    Dim sourceFolder As String
    Dim files() As String
    Dim i As Long
    Dim successCount As Long, failCount As Long
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
              "ソースフォルダ:" & vbCrLf & sourceFolder & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbQuestion, "インストール確認") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    files = Split(MODULES_LIST, "|")

    ' --- STEP 1: 既存の標準モジュールを削除 ---
    logText = "=== 削除処理 ===" & vbCrLf
    For i = 0 To UBound(files)
        Dim moduleName As String
        moduleName = GetModuleNameFromFileName(files(i))

        ' 自分自身は削除しない
        If moduleName <> "VBAインストーラー" Then
            If RemoveModule(moduleName) Then
                logText = logText & "  削除: " & moduleName & vbCrLf
            End If
        End If
    Next i

    ' --- STEP 2: .basファイルをインポート ---
    logText = logText & vbCrLf & "=== インポート処理 ===" & vbCrLf
    For i = 0 To UBound(files)
        Dim filePath As String
        filePath = sourceFolder & files(i)
        Dim modName As String
        modName = GetModuleNameFromFileName(files(i))

        If modName = "VBAインストーラー" Then
            logText = logText & "  - スキップ(自身): " & files(i) & vbCrLf
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

    ' --- STEP 3: ThisWorkbookの更新（.clsがある場合のみ） ---
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

    MsgBox "インストールが完了しました。" & vbCrLf & vbCrLf & _
           "成功: " & successCount & " モジュール" & vbCrLf & _
           "失敗: " & failCount & " モジュール" & vbCrLf & vbCrLf & _
           "詳細:" & vbCrLf & logText, _
           vbInformation, "インストール結果"
End Sub

'================================================================================
' ThisWorkbookコードを更新（.clsファイルからコード本体を取り出して挿入）
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

Private Function CheckVBEAccess() As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject.VBComponents
    CheckVBEAccess = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function GetModuleNameFromFileName(ByVal fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetModuleNameFromFileName = Left(fileName, dotPos - 1)
    Else
        GetModuleNameFromFileName = fileName
    End If
End Function

Private Function RemoveModule(ByVal moduleName As String) As Boolean
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Not comp Is Nothing Then
        If comp.Type = 3 Then  ' vbext_ct_MSForm は保護
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
