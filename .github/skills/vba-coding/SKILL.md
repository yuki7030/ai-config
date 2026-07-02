---
name: vba-coding
description: VBA(Excel/Access マクロ、.bas/.cls/.frm)のコード新規作成・修正・リファクタ時に必ず使用。Doxygen準拠コメント規約と VBA コーディング規約を定義。
---

# VBA コーディング規約

## 基本
- 全モジュール先頭に `Option Explicit`。
- 変数は型を明示宣言。`Variant` は必要時のみ。
- 命名: プロシージャ=PascalCase、ローカル変数=camelCase、定数=UPPER_SNAKE。
- エラー処理: 公開プロシージャには `On Error GoTo ErrHandler` を実装し、後始末(Close/Set Nothing)を保証。
- 画面更新抑止等(ScreenUpdating/Calculation)は必ず元に戻す。

## Doxygenヘッダコメント(必須・`'!` 形式)
※ Doxygen は VBA 非対応のためフィルタ(例: vbfilter)併用を前提に `'!` を使用。

```vb
'! @brief   売上データを集計しシートへ出力する
'! @param   wsSrc  入力元ワークシート
'! @param   dtFrom 集計開始日
'! @return  出力した行数。異常時は -1
'! @details 日付範囲外の行は無視する
Public Function AggregateSales(ByVal wsSrc As Worksheet, ByVal dtFrom As Date) As Long
    On Error GoTo ErrHandler
    ' 集計対象範囲を最終行から動的に決定(固定範囲だと行追加に追随できないため)
    ...
    Exit Function
ErrHandler:
    AggregateSales = -1
End Function
```

## 処理コメント
- ループ・分岐・API呼出等、意図が自明でない箇所に日本語で「なぜ」を記述。
- 自明なコード(`i = i + 1 ' iに1を足す` 等)へのコメントは禁止。
