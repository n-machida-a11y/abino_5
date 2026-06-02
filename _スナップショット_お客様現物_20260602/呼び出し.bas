' 標準モジュール
' 全フォーム・モジュールから呼び出せる共通ユーティリティ関数を定義する。

Option Explicit

'================================================================================
' [入口] 呼び出し（共通：ボタン → フォーム/マクロ呼び出し）
'--------------------------------------------------------------------------------
' 【何のファイル？】
'   工事名称選択や再登録など、各種フォームの起動用ラッパーマクロ。
'   ボタンから直接マクロ名を指定するだけで該当フォームが開くようにしている。
'
' 【しくみ】
'   - フォーム間で「選択された工事名・担当者」を受け渡すための
'     セル記憶ヘルパー（cellAddr に値を保存）も含む
'================================================================================


Sub StartSaitourokuProcess()
    Dim frmSelect As New 工事名称選択
    Dim frmSaitorokui As 再登録

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler_Saitouroku

    frmSelect.Show

    If Not frmSelect.Cancelled Then
        Set frmSaitorokui = New 再登録
        frmSaitorokui.SearchedKoujiName = frmSelect.selectedKoujiName
        frmSaitorokui.SelectedTantousha = frmSelect.SelectedTantousha
        frmSaitorokui.Show
    End If

    Unload frmSelect

Finalize_Saitouroku:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Set frmSelect = Nothing
    Set frmSaitorokui = Nothing
    Exit Sub

ErrorHandler_Saitouroku:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Finalize_Saitouroku
End Sub

Sub StartIraishoProcess()
    Dim frmSelect As New 工事名称選択
    Dim frmIraisho As 依頼書作成

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler_Iraisho

    frmSelect.Show

    If Not frmSelect.Cancelled Then
        Set frmIraisho = New 依頼書作成
        frmIraisho.SetupAndShow frmSelect.selectedKoujiName, frmSelect.SelectedTantousha
    End If

    Unload frmSelect

Finalize_Iraisho:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Set frmSelect = Nothing
    Set frmIraisho = Nothing
    Exit Sub

ErrorHandler_Iraisho:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Finalize_Iraisho
End Sub

'--------------------------------------------------------------------------------
' 共通ユーティリティ関数
'--------------------------------------------------------------------------------

' 外部マスターファイルのパスを返す。
' IS_TEST_MODE が True なら TEST_FILE_PATH、False なら「入力フォーム」シートから読み取る。
' 各フォーム・モジュールで同じ If IS_TEST_MODE ... パターンが重複していたため、ここに一元化。
Public Function GetTargetFilePath() As String
    If IS_TEST_MODE Then
        GetTargetFilePath = TEST_FILE_PATH
    Else
        GetTargetFilePath = Trim(CStr(ThisWorkbook.Sheets("入力フォーム").Range(PATH_CELL).Value))
    End If
End Function

' 指定したブック内に特定名称のシートが存在するかチェックする。
Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' 値が日付なら "yyyy/mm/dd" 形式の文字列を返す。日付でなければ空文字を返す。
Public Function FormatIfDate(ByVal Value As Variant) As String
    If IsDate(Value) Then
        FormatIfDate = Format(Value, "yyyy/mm/dd")
    Else
        FormatIfDate = ""
    End If
End Function

'================================================================================
' シート保護ヘルパー（複数人利用対応）
'================================================================================

' パスワード付きでシート保護を解除する。
' 1. まず設定の SHEET_PASSWORD で試行
' 2. 失敗してinteractiveがTrueならユーザーにパスワードを聞く（InputBox）
' 3. それでも失敗したら、呼び出し元のエラーハンドラに明確なメッセージで伝える
'
' 引数 interactive:
'   True  (既定)  : 解除失敗時にユーザーにパスワード入力を求める（通常の業務操作向け）
'   False         : 解除失敗時は即座にErr.Raise。
'                   InputBoxを出したくないバッチ処理（Workbook_Open等）で指定する
Public Sub SafeUnprotect(ByVal ws As Worksheet, Optional ByVal interactive As Boolean = True)
    If ws Is Nothing Then Exit Sub

    ' 1回目: 設定パスワードで試行
    On Error Resume Next
    ws.Unprotect Password:=SHEET_PASSWORD
    Err.Clear
    On Error GoTo 0

    If Not ws.ProtectContents Then Exit Sub

    ' 非対話モード: ここで即座に失敗を通知
    If Not interactive Then
        Err.Raise 1004, "SafeUnprotect", _
            "シート「" & ws.name & "」の保護解除に失敗しました。" & vbCrLf & _
            "パスワード「" & SHEET_PASSWORD & "」と一致しません。"
        Exit Sub
    End If

    ' 2回目: ユーザーにパスワードを聞いて再試行
    Dim userPwd As String
    userPwd = InputBox( _
        "シート「" & ws.name & "」の保護解除に失敗しました。" & vbCrLf & _
        "設定パスワード「" & SHEET_PASSWORD & "」では開けません。" & vbCrLf & vbCrLf & _
        "このシートのパスワードを入力してください：" & vbCrLf & _
        "（空欄/キャンセルで処理を中止します）", _
        "シート保護パスワード入力")

    If userPwd = "" Then
        Err.Raise 1004, "SafeUnprotect", _
            "シート「" & ws.name & "」の保護解除がキャンセルされました。" & vbCrLf & _
            "処理を中止します。"
        Exit Sub
    End If

    On Error Resume Next
    ws.Unprotect Password:=userPwd
    Err.Clear
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise 1004, "SafeUnprotect", _
            "シート「" & ws.name & "」の保護解除に失敗しました。" & vbCrLf & _
            "入力されたパスワードが正しくありません。" & vbCrLf & vbCrLf & _
            "【対処方法】" & vbCrLf & _
            " 1. Excelの「校閲」→「シート保護の解除」で手動解除してから再実行" & vbCrLf & _
            " 2. または、設定用コード.bas の SHEET_PASSWORD を実際のパスワードに修正"
    End If
End Sub

' データシート保護（フィルタ・並べ替えは許可、編集は不可）
' 工事番号一覧、依頼履歴、ローカルコピー用
Public Sub SafeProtectData(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=SHEET_PASSWORD, _
               UserInterfaceOnly:=False, _
               AllowFiltering:=True, _
               AllowSorting:=True, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True
    On Error GoTo 0
End Sub

' 完全保護（フィルタも編集も不可）
' 請求書提出依頼書、マスタシート用
Public Sub SafeProtectFull(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=SHEET_PASSWORD, _
               UserInterfaceOnly:=False, _
               AllowFiltering:=False, _
               AllowSorting:=False, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True
    On Error GoTo 0
End Sub

' シートのフィルタ状態をクリアする（全行表示に戻す＋AutoFilter解除）
Public Sub ClearAllFilters(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ' フィルタ適用中なら全行表示に戻す
    If ws.FilterMode Then ws.ShowAllData
    ' AutoFilterのドロップダウン矢印も消す
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
End Sub

'================================================================================
' PDF作成用のNamed Range登録ヘルパー
'   「請求書提出依頼書」シート上の、PDFファイル名に使うセル位置を
'   Named Range に記憶させる。PDF作成.bas はこの Named Range を参照することで
'   レイアウト変更（セル移動）に自動追従できる。
'
' 引数:
'   wb       : 対象ブック（通常は ThisWorkbook）
'   ws       : 対象シート（通常は 請求書提出依頼書）
'   nameId   : 登録する Named Range の名前（例: "PDF_請求宛名"）
'   cellAddr : 記憶するセル番地（例: "F5"）。空文字なら何もしない
'================================================================================
Public Sub RegisterSheetCellName(ByVal wb As Workbook, ByVal ws As Worksheet, _
                                 ByVal nameId As String, ByVal cellAddr As String)
    If wb Is Nothing Or ws Is Nothing Then Exit Sub
    If Trim(cellAddr) = "" Or Trim(nameId) = "" Then Exit Sub

    ' 既存の同名Named Rangeを削除（上書き）
    On Error Resume Next
    wb.Names(nameId).Delete
    Err.Clear
    On Error GoTo 0

    ' 新規登録
    On Error Resume Next
    wb.Names.Add name:=nameId, RefersTo:="='" & ws.name & "'!" & cellAddr
    On Error GoTo 0
End Sub


-------------------------------------------------------------------------------
