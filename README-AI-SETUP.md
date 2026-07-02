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

## Doxygenヘッダ自動検査(Hooks + CI)
3層で規約を機械的に保証する(指示文より確実):
| 層 | 設定ファイル | 動作 |
|---|---|---|
| Claude Code | .claude/settings.json | 編集直後に検査。違反はAIへ差し戻し自動修正させる(exit 2) |
| Copilot (CLI/coding agent/VS Code) | .github/hooks/doxygen.json | 編集直後に検査結果を通知 |
| CI (最終ゲート) | .github/workflows/doxygen-check.yml | PR/push時に全ファイル検査。人間のコミットも対象 |

検査本体は scripts/check_doxygen.py 1本(要 Python 3.8+)。
- C#: public/protected/internal のクラス・メソッド等に `///` ヘッダ必須
- VBA: Public Sub/Function/Property に `'!` ヘッダ必須、Option Explicit 必須
- 手動実行: `python scripts/check_doxygen.py --scan .`
- Windows で python3 コマンドが無い場合は設定内の python3 を python に読み替え(Claude Code側はフォールバック記述済み)
