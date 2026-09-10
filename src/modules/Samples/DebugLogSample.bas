Attribute VB_Name = "DebugLogSample"
'! @file    DebugLogSample.bas
'! @brief   XlflowDebug.Log の出力が VS Code のターミナルへ届くことを示すサンプル
'! @details Debug.Print はイミディエイトウィンドウにしか出ないため、xlflow 経由の
'!          実行では観測できない。XlflowDebug.Log は同じ内容を名前付きパイプで
'!          xlflow へ送り、stderr へのストリームと JSON の debug.events に載せる。
Option Explicit

'@Folder("Samples")

'* @brief   サンプル全体のエントリポイント
'* @details 入口ログ → 変数状態ログ → エラーハンドラでの Erl ログ、の3パターンを
'*          1回の実行でまとめて出力する。
Public Sub Run()
    XlflowDebug.Log "=== DebugLogSample 開始 ==="

    Call LogVariableState
    Call LogErrorLocation

    XlflowDebug.Log "=== DebugLogSample 終了 ==="
End Sub

'* @brief   変数の状態をターミナルへ出力する例
'* @details Log は ParamArray を取り、数値・真偽値・オブジェクトを自動で文字列化する。
'*          連結せず引数を並べるだけでよい。
Private Sub LogVariableState()
    Dim rowCount As Long
    Dim targetPath As String
    Dim isDryRun As Boolean

    rowCount = 42
    targetPath = "C:\temp\report.xlsx"
    isDryRun = True

    XlflowDebug.Log "LogVariableState に入りました"
    XlflowDebug.Log "rowCount=", rowCount
    XlflowDebug.Log "targetPath=", targetPath
    XlflowDebug.Log "isDryRun=", isDryRun
    XlflowDebug.Log "workbook=", ThisWorkbook.Name
End Sub

'* @brief   実行時エラーの発生行を Erl で特定する例
'* @details 行番号を振った文だけが Erl に反映される。行番号は
'*          `xlflow fmt --line-numbers add --write` で機械的に付与でき、
'*          調査後は同 remove で撤去する。
Private Sub LogErrorLocation()
    Dim total As Double
    Dim errNumber As Long
    Dim errDescription As String
    Dim errLine As Long

10      On Error GoTo ErrHandler

    ' 0除算をわざと起こし、Erl が 30 を返すことを確認するためのサンプル
30      total = 1 / 0

40      Exit Sub

ErrHandler:
    ' XlflowDebug.Log は内部で On Error を使うため、呼ぶ前に Err/Erl を退避する
    ' (先にログを出すと Erl が 0 に、Description が空になる)
    errNumber = Err.Number
    errDescription = Err.Description
    errLine = Erl

    XlflowDebug.Log "Err.Number=", errNumber
    XlflowDebug.Log "Err.Description=", errDescription
    XlflowDebug.Log "Erl=", errLine
    XlflowDebug.Log "→ Erl の値(", errLine, ")が実際に落ちた行番号"
End Sub
