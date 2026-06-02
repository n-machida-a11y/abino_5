Option Explicit

'================================================================================
' [PDF] PDF作成（依頼書を PDF として保存）
'--------------------------------------------------------------------------------
' 【何のファイル？】
'   「請求書提出依頼書」シートを PDF として書き出す処理。
'   ファイル名は「請求宛名_工事名称.pdf」のような形で自動生成される。
'
' 【しくみ】
'   - ファイル名の元になるセルは Named Range「PDF_請求宛名」「PDF_工事名称」で
'     依頼書作成フォームから事前登録される（セル番地変更に追従できる）
'================================================================================


' ファイル名に使えない文字を置換する補助的な関数
Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim illegalChars As String, i As Long
    illegalChars = "\/:*?""<>|"
    SanitizeFileName = fileName
    For i = 1 To Len(illegalChars)
        SanitizeFileName = Replace(SanitizeFileName, Mid(illegalChars, i, 1), "_")
    Next i
End Function

Public Sub SaveRequestFormAsPDF()
    Dim wsRequest As Worksheet
    Dim saveFolder As String
    Dim pdfFileName As String
    Dim fullPath As String
    Dim recipient As String
    Dim KoujiName As String
    
    '--- ① エクスポート対象のシートを指定 ---
    On Error Resume Next
    Set wsRequest = ThisWorkbook.Sheets("請求書提出依頼書")
    On Error GoTo 0
    
    If wsRequest Is Nothing Then
        MsgBox "「請求書提出依頼書」シートが見つかりません。", vbCritical
        Exit Sub
    End If
    
    '--- 保存先フォルダ ---
    ' USERPROFILE が実際のドライブと異なる環境に対応するため
    ' Z:\Users\ユーザー名\Downloads を優先して使用する
      Dim userName As String
      userName = Environ("USERNAME")
      If Dir("Z:\Users\" & userName & "\Downloads", vbDirectory) <> "" Then
          saveFolder = "Z:\Users\" & userName & "\Downloads"
      Else
          saveFolder = Environ("USERPROFILE") & "\Downloads"
      End If
    '--- ③ PDFのファイル名をシートのセルから作成 ---
    '   優先度1: 依頼書作成フォームが登録した Named Range
    '   優先度2: ハードコードの既定セル（F5, F8）
    '   Named Range が使えない場合も既定セル参照でPDF自体は発行する
    Dim nmRecipient As name, nmKoujiName As name
    Dim refInfo As String
    On Error Resume Next
    Set nmRecipient = ThisWorkbook.Names("PDF_請求宛名")
    Set nmKoujiName = ThisWorkbook.Names("PDF_工事名称")
    Err.Clear
    On Error GoTo 0

    If Not nmRecipient Is Nothing Then
        On Error Resume Next
        recipient = nmRecipient.RefersToRange.Value
        On Error GoTo 0
    End If
    If Not nmKoujiName Is Nothing Then
        On Error Resume Next
        KoujiName = nmKoujiName.RefersToRange.Value
        On Error GoTo 0
    End If

    ' フォールバック: Named Range で取れなかった場合は既定セル（F5, F8）を読む
    If Trim(recipient) = "" Then recipient = CStr(wsRequest.Range("F5").Value)
    If Trim(KoujiName) = "" Then KoujiName = CStr(wsRequest.Range("F8").Value)

    If Trim(recipient) = "" Or Trim(KoujiName) = "" Then
        ' 診断情報を組み立てる
        Dim diagMsg As String
        diagMsg = "【診断情報】" & vbCrLf
        If nmRecipient Is Nothing Then
            diagMsg = diagMsg & " PDF_請求宛名 : Named Range 未登録" & vbCrLf
        Else
            diagMsg = diagMsg & " PDF_請求宛名 : " & nmRecipient.RefersTo & " → 値=[" & recipient & "]" & vbCrLf
        End If
        If nmKoujiName Is Nothing Then
            diagMsg = diagMsg & " PDF_工事名称 : Named Range 未登録" & vbCrLf
        Else
            diagMsg = diagMsg & " PDF_工事名称 : " & nmKoujiName.RefersTo & " → 値=[" & KoujiName & "]" & vbCrLf
        End If
        diagMsg = diagMsg & " 既定セル F5 : [" & wsRequest.Range("F5").Value & "]" & vbCrLf
        diagMsg = diagMsg & " 既定セル F8 : [" & wsRequest.Range("F8").Value & "]" & vbCrLf

        MsgBox "PDFファイル名の生成に必要な情報が見つかりません。" & vbCrLf & _
               "依頼書作成フォームから書き込みを実行してから再度お試しください。" & vbCrLf & vbCrLf & _
               diagMsg, _
               vbExclamation, "PDF作成"
        Exit Sub
    End If
    
    pdfFileName = recipient & "_" & KoujiName & "_" & Format(Now, "yyyymmdd")
    pdfFileName = SanitizeFileName(pdfFileName)
    fullPath = saveFolder & "\" & pdfFileName & ".pdf"

    '--- ④ PDFとしてエクスポート ---
    Dim exportError As String
    
    wsRequest.Activate
    
    Application.ScreenUpdating = True
    On Error Resume Next
    wsRequest.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    If Err.Number <> 0 Then exportError = Err.Description
    On Error GoTo 0

    '--- ⑤ ファイルが実際に存在するか最終確認 ---
    If Dir(fullPath) <> "" Then
        ' 修正箇所: 保存先を示すメッセージを変更
        MsgBox "PDFをダウンロードフォルダに保存しました。"
    Else
        MsgBox "PDFの作成に失敗しました。" & vbCrLf & "エラー内容: " & exportError, vbCritical
    End If
    
End Sub


-------------------------------------------------------------------------------
