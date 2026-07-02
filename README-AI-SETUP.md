# AI設定ファイル 導入手順

## 配置
本一式をリポジトリのルートに展開する。

## 構成と役割
| パス | 対象 | 役割 |
|---|---|---|
| AGENTS.md | 両方 | 共通指示の唯一のソース |
| CLAUDE.md | Claude Code | AGENTS.md をインポートするだけの入口 |
| .github/copilot-instructions.md | Copilot | AGENTS.md への参照のみ |
| .github/instructions/*.instructions.md | Copilot | 拡張子別の自動適用ルール |
| .github/skills/*/SKILL.md | 両方 | 作業手順・規約本体(遅延ロードで低トークン) |
| .github/agents/*.agent.md | Copilot | 専任エージェント |
| .github/prompts/*.prompt.md | Copilot | /spec /review コマンド |
| .claude/agents/*.md | Claude Code | 専任サブエージェント |
| .claude/commands/*.md | Claude Code | /spec /review コマンド |
| docs/spec/_template.md | 両方 | 仕様書テンプレート |

## Claude Code でスキルを共有する(必須)
スキル本体は .github/skills/ に一元化。Claude Code からはリンクで共有:
- macOS/Linux: `ln -s ../.github/skills .claude/skills`
- Windows(管理者不要): `mklink /J .claude\skills .github\skills`
リンク不可の環境ではディレクトリをコピーして同期する。

## 運用フロー
1. `/spec <要求>` → 仕様書起案 → 人が承認
2. 実装依頼(vba-developer / csharp-developer)
3. `/review` → 指摘対応 → コミット

## トークン節約の仕組み
- 常駐するのは AGENTS.md(短文)とスキルの説明文のみ。
- 規約本体(SKILL.md)は関連タスク時のみロード。
- 指示の重複を排し、参照(リンク)で一元管理。

## 剪定運用(月1回・重要)
設定ファイルは毎リクエストでコンテキストに乗り続けるコスト。以下を定期確認:
1. AIが繰り返し無視したルール → 表現を強める or スキル/Hooksへ移す
2. 古くなったルール → 削除
3. AGENTS.md が40行超 → スキルへ分離
4. 「なぜ」が不明なルール → 理由を1句添える(遵守率・応用力が向上)
