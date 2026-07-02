---
name: csharp-coding
description: C#(.cs)のコード新規作成・修正・リファクタ時に必ず使用。Doxygen準拠コメント規約と C# コーディング規約を定義。
---

# C# コーディング規約

## 基本
- .NET 標準命名規則(型/メソッド=PascalCase、ローカル/引数=camelCase、プライベートフィールド=_camelCase)。
- Nullable 有効前提。null 許容は `?` で明示。
- 例外は握りつぶさない。catch するなら記録または再スロー。
- `using` / `await using` でリソース解放。非同期は async/await、`.Result`/`.Wait()` 禁止。
- マジックナンバー禁止。定数または enum 化。

## Doxygenヘッダコメント(必須・`///` + @コマンド形式)
```csharp
/// @brief     指定期間の売上を集計する
/// @param     from 集計開始日(この日を含む)
/// @param     to   集計終了日(この日を含む)
/// @return    集計結果。該当なしの場合は空リスト
/// @exception ArgumentException from > to の場合
public IReadOnlyList<SalesSummary> Aggregate(DateOnly from, DateOnly to)
{
    if (from > to) throw new ArgumentException("期間指定が不正");
    // DBではなくメモリ上で集計する(件数が高々数千件でありRTを優先するため)
    ...
}
```
- クラス・公開プロパティにも `/// @brief` を付与。

## 処理コメント
- 意図が自明でないブロックに日本語で「なぜ」を記述。自明なコメントは禁止。
