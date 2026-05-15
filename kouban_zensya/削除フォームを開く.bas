Attribute VB_Name = "削除フォームを開く"
Option Explicit

' [FIX] 元のコードは Exit Sub なしに ErrorHandler: へ流れ落ちる構造になっており、
'       正常終了後も Err.Number チェックが走っていた（実害はないが意図が不明瞭）。
'       また、エラー発生時にフォームが Unload されないリソースリークがあった。
'       Exit Sub を追加し、エラー時の Unload 処理を ErrorHandler 内に移動した。
Sub StartSakujoProcess()
    Dim frmSakujo As New 削除フォーム

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    frmSakujo.Show
    Unload frmSakujo
    Application.ScreenUpdating = True

    Exit Sub  ' 正常終了時はここで終わり。以下の ErrorHandler へは流れない。

ErrorHandler:
    ' エラー時もフォームを確実に閉じる
    On Error Resume Next
    Unload frmSakujo
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

