# vba-tool

Excel VBA(+ 将来的に C#)で業務ツールを開発するリポジトリ。
VBA 開発は [xlflow](docs/README-XLFLOW.md) で行い、VBE は使用しない。
AI エージェント(Claude Code / GitHub Copilot)前提の開発フローを備える。

## 主な機能(VBA ワークシート関数)

| 関数 | モジュール | 概要 |
|---|---|---|
|---|---|---|

## 技術スタック

- VBA: Excel VBA 7.1(Office 2019+ / 32bit)
- C#: .NET 8 / C# 12 / Nullable 有効(該当タスク発生時)

## ディレクトリ構成

| パス | 役割 |
|---|---|
| [src/](src/) | VBA ソース(`modules/` = 標準モジュール、`classes/` = クラス、`workbook/` = ブック/シートモジュール)。xlflow の編集対象 |
| [src/modules/<機能>/](src/modules/) | 機能ごとにフォルダ分割(`@Folder` 注釈と対応)。本体とテスト(`*Tests.bas`)を同梱 |
| `build/vba-tool.xlsm` | ビルド対象ブック(`.gitignore` 済み。xlflow が push/save する実体) |
| [docs/](docs/) | プロジェクト知識(用語・スキーマ・業務ルール・仕様・監査記録) |
| [.github/skills/](.github/skills/) | 作業手順・規約の本体(SKILL.md。遅延ロード) |
| [.claude/](.claude/) | Claude Code のサブエージェント・コマンド・設定 |
| [scripts/](scripts/) | 静的解析(Doxygen ヘッダ検査 / VBA lint)・指示ファイル監査 |
| [xlflow.toml](xlflow.toml) | xlflow のプロジェクト設定 |

## セットアップ

1. xlflow・VSCode 拡張・トラストセンター設定など環境構築は
   [docs/README-XLFLOW.md](docs/README-XLFLOW.md) を参照(完全自動スクリプト [scripts/xlflow-setup.ps1](scripts/xlflow-setup.ps1) あり)。
2. AI 設定ファイル(AGENTS.md / スキル / エージェント等)の導入は
   [docs/README-AI-SETUP.md](docs/README-AI-SETUP.md) を参照。

## 開発フロー(VBA / xlflow)

```powershell
xlflow session start
# src/ 配下の .bas / .cls を編集(UTF-8。CP932 変換は push 時に自動)
xlflow push --fast --session --no-save --json   # 反映(preflight で lint)
xlflow test --session --json                     # テスト実行
xlflow save --json ; xlflow session stop
```

- 詳細な手順・安全規則は [.github/skills/xlflow/SKILL.md](.github/skills/xlflow/SKILL.md)。
- `Debug.Print` 禁止 → `XlflowDebug.Log`。`MsgBox`/`InputBox` 直接使用禁止 → `XlflowUI` ラッパー。
- テストは各機能フォルダに `Test` 接頭辞の Sub で書き、`XlflowAssert` でアサーションする。

## ドキュメント

- 共通開発指示(唯一のソース): [AGENTS.md](AGENTS.md)(Claude Code は [CLAUDE.md](CLAUDE.md) 経由で参照)
- xlflow 導入ガイド: [docs/README-XLFLOW.md](docs/README-XLFLOW.md)
- AI 設定ファイル導入手順: [docs/README-AI-SETUP.md](docs/README-AI-SETUP.md)
- プロジェクト知識: [用語集](docs/glossary.md) / [スキーマ](docs/schema.md) / [業務ルール](docs/business-rules.md)
