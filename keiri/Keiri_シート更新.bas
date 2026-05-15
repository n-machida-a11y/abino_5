Option Explicit

'================================================================================
' 設定・定数定義
' IS_TEST_MODE / SHEET_IRAI_* / GetMasterPath などは Keiri_設定用コード (Config)
' モジュールで一元管理されています。このモジュールでは定義しない。
'================================================================================
' シート名用の別名（既存コード互換のため）
Public Const m_IRAI_RIREKI_SHEET_NAME As String = "依頼履歴"
Public Const m_IRAI_SEARCH_SHEET_NAME As String = "依頼検索"

'--- 依頼検索シートのセル配置定義 ---
Private Const CELL_IRAI_NO As String = "A2"       ' 検索キー
Private Const CELL_KEIRI_NO As String = "B2"      ' 請求書番号(経理用)

' ■5行目エリア（主要項目）
Private Const CELL_PROJ_NO As String = "A5"       ' 工事番号
Private Const CELL_DATE_SUBMIT As String = "B5"   ' 請求書日付
Private Const CELL_DESTINATION As String = "C5"   ' 請求宛名
Private Const CELL_PROJ_NAME As String = "D5"     ' 工事名称
Private Const CELL_AMOUNT As String = "E5"        ' 請負金額(税込)
Private Const CELL_STAFF As String = "F5"         ' 担当者

' ■7行目エリア（新レイアウト）
Private Const CELL_TERM_START As String = "A7"    ' 工期 着手
Private Const CELL_TERM_END As String = "B7"      ' 工期 完成
Private Const CELL_DATE_DELIVERY As String = "C7" ' 引渡日 (旧G9)
Private Const CELL_BANK_ACCOUNT As String = "D7"  ' 但陽信金口座指定 (旧F7)
Private Const CELL_COMMENT As String = "E7"       ' 担当者引継ぎコメント (旧F9)
Private Const CELL_DATE_CREATE As String = "F7"   ' 依頼書作成日 (旧A11)

' ■9行目エリア（要項・住所・詳細）
Private Const CELL_REQ_ITEM As String = "A9"      ' 提出要項
Private Const CELL_ENCLOSURE As String = "B9"     ' 同封物 (旧E9)
Private Const CELL_POST_CODE As String = "C9"     ' 郵便番号
Private Const CELL_ADDRESS As String = "D9"       ' 郵送先住所
Private Const CELL_MEMO As String = "E9"          ' 経理メモ (旧C11)
Private Const CELL_DATE_ISSUE As String = "F9"    ' 作成日 (旧B11)

' ■11行目エリア（管理情報）
Private Const CELL_HISTORY As String = "A11"      ' 操作履歴 (旧D11)

'--- 廃止フィールド（新レイアウトでは 依頼検索 に存在しない）---
'   CELL_REQ_OTHER     (旧 B9, 提出要項(その他))
'   CELL_TEXT_INVOICE  (旧 C7, 請求書記載文言)
'   CELL_NOTE_RECEIPT  (旧 D7, 領収書注意文)
'   CELL_NOTE_FEE      (旧 E7, 振込手数料注意文)


'================================================================================
' ボタン登録用
'================================================================================
Public Sub マクロ_検索実行()
    Call ExecuteUpdateAndSearch
End Sub

Public Sub マクロ_上書き保存()
    Dim wsRireki As Worksheet: Set wsRireki = ThisWorkbook.Sheets(m_IRAI_RIREKI_SHEET_NAME)
    Dim wsSearch As Worksheet: Set wsSearch = ThisWorkbook.Sheets(m_IRAI_SEARCH_SHEET_NAME)
    Dim targetIraiNo As String: targetIraiNo = Trim(wsSearch.Range(CELL_IRAI_NO).Value)
    Dim foundRange As Range, autoReason As String
    
    If targetIraiNo = "" Then
        MsgBox "依頼NOを指定してください。"
        Exit Sub
    End If

    ' 完全一致で検索（上書き時はIDを変えないため）
    Set foundRange = wsRireki.Columns("A").Find(What:=targetIraiNo, LookAt:=xlWhole)
    If foundRange Is Nothing Then
        MsgBox "指定の依頼NOが見つかりません。" & vbCrLf & "（新規の場合は「新規作成」ボタンを使用してください）"
        Exit Sub
    End If
    
    autoReason = GetDifferenceLog(wsRireki, wsSearch, foundRange.Row)
    
    Dim msg As String
    If autoReason = "" Then
        msg = "変更箇所が見つかりませんでした。上書きしますか？"
    Else
        msg = "以下の変更を同期して上書きしますか？" & vbCrLf & autoReason
    End If

    If MsgBox(msg, vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    Call WriteDataToRireki(wsRireki, wsSearch, foundRange.Row, False, autoReason, "")
    ' マスタ同期成功時のみローカル物理保存。失敗時は保存せずロールバック可能に。
    If SyncAllDataToMaster(targetIraiNo, False, autoReason) Then
        ThisWorkbook.Save
        MsgBox "同期完了しました。", vbInformation
    End If
End Sub

Public Sub マクロ_新規作成()
    Dim wsRireki As Worksheet: Set wsRireki = ThisWorkbook.Sheets(m_IRAI_RIREKI_SHEET_NAME)
    Dim wsSearch As Worksheet: Set wsSearch = ThisWorkbook.Sheets(m_IRAI_SEARCH_SHEET_NAME)
    Dim lastRow As Long
    Dim oldIraiNo As String: oldIraiNo = Trim(wsSearch.Range(CELL_IRAI_NO).Value)
    Dim parentID As String
    Dim newIraiNo As String
    
    ' 親IDの入力を求める（デフォルト値として現在のIDを入れる）
    ' そのままEnterで枝番作成、消してEnterで完全新規、キャンセルで中止
    parentID = InputBox("親となる依頼NOを指定してください。" & vbCrLf & _
                        "（この番号の枝番を作成します）" & vbCrLf & vbCrLf & _
                        "※完全新規（連番）の場合は、ここを空欄にしてOKを押してください。", _
                        "新規・枝番発行", oldIraiNo)
                        
    If StrPtr(parentID) = 0 Then Exit Sub ' キャンセルボタンが押された場合
    
    If parentID <> "" Then
        ' --- 枝番発行（デフォルト） ---
        newIraiNo = GetNextBranchID(wsRireki, parentID)
    Else
        ' --- 完全新規（連番） ---
        newIraiNo = GetNextIntegerID(wsRireki)
    End If
    
    ' 新規IDセット
    wsSearch.Range(CELL_IRAI_NO).Value = newIraiNo
    ' 経理番号は空にする
    wsSearch.Range(CELL_KEIRI_NO).Value = ""
    
    ' 履歴シートへの行追加
    lastRow = wsRireki.Cells(wsRireki.Rows.Count, "A").End(xlUp).Row
    Dim nextRow As Long: nextRow = lastRow + 1
    wsRireki.Cells(nextRow, "A").Value = newIraiNo ' 文字列としてセットされるよう注意
    
    Dim logMsg As String
    logMsg = "新規発行 (ID:" & newIraiNo & ")"
    If parentID <> "" Then
        logMsg = logMsg & " (親ID:" & parentID & "からの枝番)"
    End If
    
    Call WriteDataToRireki(wsRireki, wsSearch, nextRow, True, "", oldIraiNo)
    ' 新ID告知を同期の前に出す。失敗時は「新規発行→同期エラー」の順で表示される。
    MsgBox "新規発行: " & newIraiNo, vbInformation
    ' マスタ同期成功時のみローカル物理保存。
    If SyncAllDataToMaster(newIraiNo, True, logMsg) Then
        ThisWorkbook.Save
    End If
End Sub

'================================================================================
' 内部ロジック（ID生成・検索・同期）
'================================================================================

Private Function GetNextIntegerID(ws As Worksheet) As String
    Dim lastRow As Long, i As Long
    Dim maxVal As Long, v As Variant
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    maxVal = 0
    For i = 3 To lastRow
        v = ws.Cells(i, 1).Value
        If IsNumeric(v) Then
            If InStr(v, "-") = 0 Then
                If CLng(v) > maxVal Then maxVal = CLng(v)
            End If
        End If
    Next i
    GetNextIntegerID = CStr(maxVal + 1)
End Function

Private Function GetNextBranchID(ws As Worksheet, parentID As String) As String
    Dim lastRow As Long, i As Long
    Dim maxBranch As Long, currentBranch As Long
    Dim v As String
    Dim parts As Variant
    
    If InStr(parentID, "-") > 0 Then
        parentID = Split(parentID, "-")(0)
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    maxBranch = 0
    
    For i = 3 To lastRow
        v = CStr(ws.Cells(i, 1).Value)
        If v Like parentID & "-*" Then
            parts = Split(v, "-")
            If UBound(parts) >= 1 Then
                If IsNumeric(parts(1)) Then
                    currentBranch = CLng(parts(1))
                    If currentBranch > maxBranch Then maxBranch = currentBranch
                End If
            End If
        End If
    Next i
    
    GetNextBranchID = parentID & "-" & (maxBranch + 1)
End Function

Private Sub WriteDataToRireki(wsRireki As Worksheet, wsSearch As Worksheet, targetRow As Long, _
                              isNew As Boolean, autoReason As String, oldIraiNo As String)
    With wsRireki
        .Cells(targetRow, 2).Value = wsSearch.Range(CELL_KEIRI_NO).Value
        .Cells(targetRow, 3).Value = wsSearch.Range(CELL_DATE_ISSUE).Value
        .Cells(targetRow, 4).Value = wsSearch.Range(CELL_STAFF).Value
        .Cells(targetRow, 5).Value = wsSearch.Range(CELL_PROJ_NO).Value
        .Cells(targetRow, 6).Value = wsSearch.Range(CELL_PROJ_NAME).Value
        .Cells(targetRow, 7).Value = wsSearch.Range(CELL_TERM_START).Value
        .Cells(targetRow, 8).Value = wsSearch.Range(CELL_TERM_END).Value
        .Cells(targetRow, 9).Value = wsSearch.Range(CELL_AMOUNT).Value
        
        .Cells(targetRow, 10).Value = wsSearch.Range(CELL_DATE_CREATE).Value
        .Cells(targetRow, 11).Value = wsSearch.Range(CELL_REQ_ITEM).Value
        ' col 12 (提出要項その他) 廃止につき書き込まない
        .Cells(targetRow, 13).Value = wsSearch.Range(CELL_ENCLOSURE).Value
        .Cells(targetRow, 14).Value = wsSearch.Range(CELL_DATE_SUBMIT).Value
        .Cells(targetRow, 15).Value = wsSearch.Range(CELL_DATE_DELIVERY).Value
        .Cells(targetRow, 16).Value = wsSearch.Range(CELL_DESTINATION).Value
        .Cells(targetRow, 17).Value = wsSearch.Range(CELL_POST_CODE).Value
        .Cells(targetRow, 18).Value = wsSearch.Range(CELL_ADDRESS).Value
        
        ' col 19 (請求書記載文言) 廃止につき書き込まない
        .Cells(targetRow, 20).Value = wsSearch.Range(CELL_COMMENT).Value
        ' col 21 (領収書注意文) 廃止につき書き込まない
        ' col 22 (振込手数料注意文) 廃止につき書き込まない
        .Cells(targetRow, 23).Value = wsSearch.Range(CELL_BANK_ACCOUNT).Value
    End With
    
    Dim uiHistory As String: uiHistory = wsSearch.Range(CELL_HISTORY).Value
    Dim uiMemo As String: uiMemo = wsSearch.Range(CELL_MEMO).Value
    Dim timeStamp As String: timeStamp = Format(Now, "yyyy/mm/dd hh:mm")
    Dim entry As String: entry = ""
    
    If isNew Then
        entry = "【新規発行】 " & timeStamp & IIf(oldIraiNo <> "", " (元:No." & oldIraiNo & "から参照作成)", "")
    ElseIf autoReason <> "" Then
        entry = "【更新】 " & timeStamp & vbCrLf & autoReason
    End If
    
    If entry <> "" Then
        If uiHistory = "" Then
            wsRireki.Cells(targetRow, 24).Value = entry
        Else
            wsRireki.Cells(targetRow, 24).Value = uiHistory & vbCrLf & "---" & vbCrLf & entry
        End If
    Else
        wsRireki.Cells(targetRow, 24).Value = uiHistory
    End If
    
    wsRireki.Cells(targetRow, 25).Value = uiMemo
    
    wsSearch.Range(CELL_HISTORY).Value = wsRireki.Cells(targetRow, 24).Value
    wsSearch.Range(CELL_MEMO).Value = wsRireki.Cells(targetRow, 25).Value
    ' ThisWorkbook.Save は呼び出し元で「マスタ同期成功時のみ」実行する
    ' （先に物理保存するとマスタ同期失敗時にローカルだけ新値で残って不整合になる）
End Sub

Private Function GetDifferenceLog(wsRireki As Worksheet, wsSearch As Worksheet, targetRow As Long) As String
    Dim log As String: log = ""
    Dim checkItems, i As Long
    
    checkItems = Array( _
        Array("担当者", CELL_STAFF, 4), _
        Array("工事番号", CELL_PROJ_NO, 5), _
        Array("工事名称", CELL_PROJ_NAME, 6), _
        Array("工期着手", CELL_TERM_START, 7), _
        Array("工期完成", CELL_TERM_END, 8), _
        Array("請負金額", CELL_AMOUNT, 9), _
        Array("提出日付", CELL_DATE_SUBMIT, 14), _
        Array("宛先", CELL_DESTINATION, 16), _
        Array("発行日", CELL_DATE_ISSUE, 3), _
        Array("提出要項", CELL_REQ_ITEM, 11), _
        Array("郵便番号", CELL_POST_CODE, 17), _
        Array("住所", CELL_ADDRESS, 18) _
    )
                            
    For i = LBound(checkItems) To UBound(checkItems)
        Dim nV: nV = wsSearch.Range(checkItems(i)(1)).Value
        Dim oV: oV = wsRireki.Cells(targetRow, checkItems(i)(2)).Value
        If CStr(nV) <> CStr(oV) Then
            log = log & checkItems(i)(0) & "：" & oV & " → " & nV & vbCrLf
        End If
    Next i
    GetDifferenceLog = log
End Function

Public Function SyncAllDataToMaster(ByVal targetIraiNo As String, ByVal isNew As Boolean, ByVal logMsg As String) As Boolean
    SyncAllDataToMaster = False  ' 既定は失敗。全処理成功で True

    Dim mPath As String, wbM As Workbook, wsM As Worksheet, wsS As Worksheet
    Dim targetRow As Long
    
    Set wsS = ThisWorkbook.Sheets(m_IRAI_SEARCH_SHEET_NAME)
    mPath = GetMasterPath()
    
    ' 空パス対策: Dir("") はエラー52を投げるので事前チェック
    If Trim(mPath) = "" Then
        MsgBox "マスタファイルのパスが設定されていません。" & vbCrLf & _
               "依頼履歴シートの G1 セルにマスタファイルのパスを設定してください。", vbCritical
        Exit Function
    End If
    On Error Resume Next
    Dim dirChk As String: dirChk = Dir(mPath)
    On Error GoTo 0
    If dirChk = "" Then
        MsgBox "マスタファイルにアクセスできません。" & vbCrLf & _
               "パス: " & mPath & vbCrLf & vbCrLf & _
               "ネットワーク(Z:)の接続を確認してください。" & vbCrLf & vbCrLf & _
               "変更はローカルRAMに残っていますが物理保存されていません。" & vbCrLf & _
               "接続回復後、同じ操作を再実行してください。", vbCritical, "マスタ同期エラー"
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    On Error Resume Next
    Set wbM = Workbooks.Open(mPath)
    On Error GoTo 0
    
    If wbM Is Nothing Then
        MsgBox "マスタファイルを開けませんでした。" & vbCrLf & _
               "他のユーザーが編集中の可能性があります。", vbCritical
        Application.ScreenUpdating = True
        Exit Function
    End If
    
    Set wsM = wbM.Sheets(m_IRAI_RIREKI_SHEET_NAME)
    
    If isNew Then
        targetRow = wsM.Cells(wsM.Rows.Count, "A").End(xlUp).Row + 1
        wsM.Cells(targetRow, "A").Value = targetIraiNo
    Else
        Dim f As Range
        Set f = wsM.Columns("A").Find(targetIraiNo, LookAt:=xlWhole)
        If f Is Nothing Then
            MsgBox "マスタ内に依頼NO [" & targetIraiNo & "] が見つかりませんでした。", vbExclamation
            wbM.Close False
            Application.ScreenUpdating = True
            Exit Function
        End If
        targetRow = f.Row
    End If
    
    With wsM
        .Cells(targetRow, 2).Value = wsS.Range(CELL_KEIRI_NO).Value
        .Cells(targetRow, 3).Value = wsS.Range(CELL_DATE_ISSUE).Value
        .Cells(targetRow, 4).Value = wsS.Range(CELL_STAFF).Value
        .Cells(targetRow, 5).Value = wsS.Range(CELL_PROJ_NO).Value
        .Cells(targetRow, 6).Value = wsS.Range(CELL_PROJ_NAME).Value
        .Cells(targetRow, 7).Value = wsS.Range(CELL_TERM_START).Value
        .Cells(targetRow, 8).Value = wsS.Range(CELL_TERM_END).Value
        .Cells(targetRow, 9).Value = wsS.Range(CELL_AMOUNT).Value
        
        .Cells(targetRow, 10).Value = wsS.Range(CELL_DATE_CREATE).Value
        .Cells(targetRow, 11).Value = wsS.Range(CELL_REQ_ITEM).Value
        ' col 12 (提出要項その他) 廃止につき書き込まない
        .Cells(targetRow, 13).Value = wsS.Range(CELL_ENCLOSURE).Value
        .Cells(targetRow, 14).Value = wsS.Range(CELL_DATE_SUBMIT).Value
        .Cells(targetRow, 15).Value = wsS.Range(CELL_DATE_DELIVERY).Value
        .Cells(targetRow, 16).Value = wsS.Range(CELL_DESTINATION).Value
        .Cells(targetRow, 17).Value = wsS.Range(CELL_POST_CODE).Value
        .Cells(targetRow, 18).Value = wsS.Range(CELL_ADDRESS).Value
        
        ' col 19 (請求書記載文言) 廃止につき書き込まない
        .Cells(targetRow, 20).Value = wsS.Range(CELL_COMMENT).Value
        ' col 21 (領収書注意文) 廃止につき書き込まない
        ' col 22 (振込手数料注意文) 廃止につき書き込まない
        .Cells(targetRow, 23).Value = wsS.Range(CELL_BANK_ACCOUNT).Value
        
        .Cells(targetRow, 24).Value = wsS.Range(CELL_HISTORY).Value
        .Cells(targetRow, 25).Value = wsS.Range(CELL_MEMO).Value
    End With
    
    wbM.Save
    wbM.Close False
    Application.ScreenUpdating = True
    SyncAllDataToMaster = True  ' ここまで到達で成功
End Function

' GetMasterPath() は Keiri_設定用コード (Config モジュール) 側の
' 同名関数を使う。重複定義を避けるためこちら側は削除済み。

Function ExecuteUpdateAndSearch() As Boolean
    ExecuteUpdateAndSearch = False
    
    If UpdateKeiriRirekiSheet(False) = False Then
        Exit Function
    End If
    
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets(m_IRAI_RIREKI_SHEET_NAME)
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets(m_IRAI_SEARCH_SHEET_NAME)
    
    Dim searchKey As String
    searchKey = Trim(wsS.Range(CELL_IRAI_NO).Value)
    
    If searchKey = "" Then
        MsgBox "依頼NOを入力してください。"
        Exit Function
    End If
    
    ' --- 候補検索ロジック ---
    Dim candidates As Collection: Set candidates = New Collection
    Dim lastRow As Long: lastRow = wsR.Cells(wsR.Rows.Count, "A").End(xlUp).Row
    Dim i As Long, v As String
    Dim targetRow As Long
    
    For i = 3 To lastRow
        v = CStr(wsR.Cells(i, 1).Value)
        If v = searchKey Or v Like searchKey & "-*" Then
            candidates.Add i
        End If
    Next i
    
    If candidates.Count = 0 Then
        Call ClearSearchSheet(wsS)
        MsgBox "該当するデータが見つかりません。", vbInformation
        Exit Function
        
    ElseIf candidates.Count = 1 Then
        targetRow = candidates(1)
        
    Else
        ' 複数候補があるので、ユーザーにどれを選ぶか聞く
        ' UserForm 依存を避けるため、標準の InputBox で候補一覧を表示して番号で選んでもらう方式
        Dim prompt As String, idx As Long
        prompt = "複数の依頼NO候補が見つかりました。" & vbCrLf & _
                 "選択したい番号を入力してOKを押してください:" & vbCrLf & vbCrLf
        For idx = 1 To candidates.Count
            prompt = prompt & idx & ": " & CStr(wsR.Cells(candidates(idx), 1).Value)
            Dim proj As String: proj = CStr(wsR.Cells(candidates(idx), 5).Value) ' 工事番号
            Dim nam As String: nam = CStr(wsR.Cells(candidates(idx), 6).Value)  ' 工事名称
            If proj <> "" Then prompt = prompt & " [" & proj & "]"
            If nam <> "" Then prompt = prompt & " " & nam
            prompt = prompt & vbCrLf
        Next idx

        Dim picked As String
        picked = InputBox(prompt, "依頼NOの選択", "1")
        If picked = "" Then Exit Function  ' キャンセル or 空欄
        If Not IsNumeric(picked) Then
            MsgBox "番号（1〜" & candidates.Count & "）を入力してください。", vbExclamation
            Exit Function
        End If
        Dim pickIdx As Long: pickIdx = CLng(picked)
        If pickIdx < 1 Or pickIdx > candidates.Count Then
            MsgBox "範囲外です（1〜" & candidates.Count & "）。", vbExclamation
            Exit Function
        End If
        targetRow = candidates(pickIdx)
        wsS.Range(CELL_IRAI_NO).Value = CStr(wsR.Cells(targetRow, 1).Value)
    End If
    
    Call TransferRowToSearchSheet(wsR, wsS, targetRow)
    ExecuteUpdateAndSearch = True
End Function

Private Sub ClearSearchSheet(wsS As Worksheet)
    wsS.Range(CELL_KEIRI_NO & "," & CELL_PROJ_NO & "," & CELL_PROJ_NAME & "," & CELL_AMOUNT).ClearContents
    wsS.Range(CELL_DESTINATION & "," & CELL_TERM_START & "," & CELL_TERM_END).ClearContents
    wsS.Range(CELL_BANK_ACCOUNT).ClearContents
    wsS.Range(CELL_REQ_ITEM & "," & CELL_POST_CODE & "," & CELL_ADDRESS).ClearContents
    wsS.Range(CELL_ENCLOSURE & "," & CELL_COMMENT & "," & CELL_DATE_DELIVERY).ClearContents
    wsS.Range(CELL_DATE_CREATE & "," & CELL_DATE_ISSUE).ClearContents
    wsS.Range(CELL_HISTORY & "," & CELL_MEMO).ClearContents
End Sub

Private Sub TransferRowToSearchSheet(wsR As Worksheet, wsS As Worksheet, r As Long)
    With wsS
        .Range(CELL_KEIRI_NO).Value = wsR.Cells(r, 2).Value
        .Range(CELL_DATE_ISSUE).Value = wsR.Cells(r, 3).Value
        .Range(CELL_STAFF).Value = wsR.Cells(r, 4).Value
        .Range(CELL_PROJ_NO).Value = wsR.Cells(r, 5).Value
        .Range(CELL_PROJ_NAME).Value = wsR.Cells(r, 6).Value
        .Range(CELL_TERM_START).Value = wsR.Cells(r, 7).Value
        .Range(CELL_TERM_END).Value = wsR.Cells(r, 8).Value
        
        ' ★改修：金額をセットし、￥マーク書式を適用
        ' Rangeの前にドット(.)を付与、書式を円記号に修正
        .Range(CELL_AMOUNT).Value = wsR.Cells(r, 9).Value
        .Range(CELL_AMOUNT).NumberFormatLocal = "\#,##0"
        
        .Range(CELL_DATE_CREATE).Value = wsR.Cells(r, 10).Value
        .Range(CELL_REQ_ITEM).Value = wsR.Cells(r, 11).Value
        ' col 12 (提出要項その他) 廃止につき読まない
        .Range(CELL_ENCLOSURE).Value = wsR.Cells(r, 13).Value
        .Range(CELL_DATE_SUBMIT).Value = wsR.Cells(r, 14).Value
        .Range(CELL_DATE_DELIVERY).Value = wsR.Cells(r, 15).Value
        .Range(CELL_DESTINATION).Value = wsR.Cells(r, 16).Value
        .Range(CELL_POST_CODE).Value = wsR.Cells(r, 17).Value
        .Range(CELL_ADDRESS).Value = wsR.Cells(r, 18).Value
        
        ' col 19 (請求書記載文言) 廃止につき読まない
        .Range(CELL_COMMENT).Value = wsR.Cells(r, 20).Value
        ' col 21 (領収書注意文) 廃止につき読まない
        ' col 22 (振込手数料注意文) 廃止につき読まない
        .Range(CELL_BANK_ACCOUNT).Value = wsR.Cells(r, 23).Value
        
        .Range(CELL_HISTORY).Value = wsR.Cells(r, 24).Value
        .Range(CELL_MEMO).Value = wsR.Cells(r, 25).Value
    End With
End Sub

Function UpdateKeiriRirekiSheet(Optional ByVal ShowMessage As Boolean = True) As Boolean
    Dim wbT As Workbook, wsSrc As Worksheet, wsDst As Worksheet, p As String, l As Long
    UpdateKeiriRirekiSheet = False
    Application.ScreenUpdating = False
    p = GetMasterPath()
    
    ' 空パス対策: Dir("") はエラー52(ファイル名または番号が不正)を投げるので事前チェック
    If Trim(p) = "" Then
        Application.ScreenUpdating = True
        MsgBox "マスタファイルのパスが設定されていません。" & vbCrLf & _
               "依頼履歴シートの G1 セルにマスタファイルのパスを設定してください。", vbCritical
        Exit Function
    End If

    On Error Resume Next
    Dim dirResult As String: dirResult = Dir(p)
    On Error GoTo 0
    If dirResult = "" Then
        Application.ScreenUpdating = True
        MsgBox "マスタファイルにアクセスできません。" & vbCrLf & _
               "パス: " & p, vbCritical
        Exit Function
    End If
    
    On Error Resume Next
    Set wbT = Workbooks.Open(p, ReadOnly:=True)
    On Error GoTo 0
    If wbT Is Nothing Then Exit Function
    
    Set wsSrc = wbT.Sheets(m_IRAI_RIREKI_SHEET_NAME)
    Set wsDst = ThisWorkbook.Sheets(m_IRAI_RIREKI_SHEET_NAME)
    
    wsDst.Range("A3:Z" & wsDst.Rows.Count).Clear
    l = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    If l >= 2 Then
        wsSrc.Range("A2:Z" & l).Copy wsDst.Range("A3")
    End If
    
    wbT.Close False
    If ShowMessage Then MsgBox "最新データを取得しました。"
    UpdateKeiriRirekiSheet = True
    Application.ScreenUpdating = True
End Function

