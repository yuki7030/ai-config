# 共通開発指示(Claude Code / GitHub Copilot 共用)

## 技術スタック
<!-- バージョンまで具体的に書くこと(効果:極高)。例: -->
- VBA: Excel VBA 7.1 (Office 2019+ / 32bit)
- C#: .NET 8 / C# 12 / Nullable有効

## 進め方(ALWAYS)
- 実装前に必ず「調査→計画提示→承認→実装→検証」の順で進める。いきなりコーディングしない。
- 変更が200行を超えそうな場合、着手前に規模と方針を提示して確認を取る。
- 要求が曖昧なら推測せず質問する(最大5件)。
- 新機能・仕様変更は docs/spec/ の仕様書が前提。無ければ spec-writing スキルで起案し承認を得る。

## 禁止(NEVER)
- 指示されたスコープ外のファイルを変更しない。
- 仕様書に無い機能を追加しない。
- エラーの握りつぶし(空catch / On Error Resume Next の放置)をしない。
- 秘密情報(接続文字列・パスワード)をハードコードしない。
- 無関係なリファクタ・整形をしない(差分レビューを汚すため)。

## 出力・コメント
- 回答・コメント・ドキュメントは日本語、識別子は英語。回答は要点のみ簡潔に。
- 全公開関数に Doxygen ヘッダ必須(C#: `///`+@brief等 / VBA: `'!`+@brief等)。
- 処理コメントは「なぜ」を書く(自明な「何を」コメントは禁止)。
- コミット: プレフィックス英語(feat:/fix:/docs:/chore:)+本文日本語。

## 詳細規約(タスク該当時に参照)
- 仕様書作成: .github/skills/spec-writing/
- VBA実装: .github/skills/vba-coding/
- C#実装: .github/skills/csharp-coding/
- レビュー: .github/skills/code-review/

## 検証
<!-- プロジェクトのビルド/テストコマンドを記載。例: dotnet build && dotnet test -->
