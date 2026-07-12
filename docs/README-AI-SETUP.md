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
| .github/prompts/*.prompt.md | Copilot | /spec /review /audit-instructions コマンド |
| .claude/agents/*.md | Claude Code | 専任サブエージェント |
| .claude/commands/*.md | Claude Code | /spec /review /audit-instructions コマンド |
| docs/spec/_template.md | 両方 | 仕様書テンプレート |
| scripts/audit_instructions.py ほか | 両方 | 指示ファイルの機械監査(月次自動・下記) |

## Claude Code でスキルを共有する(必須)
スキル本体は .github/skills/ に一元化。Claude Code からはリンクで共有:
- macOS/Linux: `ln -s ../.github/skills .claude/skills`
- Windows(管理者不要): `mklink /J .claude\skills .github\skills`
リンク不可の環境ではディレクトリをコピーして同期する。

## 運用フロー
### A. 自律モード(推奨): `/implement <要求 または 仕様書パス>`
仕様書起案 → **人が承認(唯一のゲート)** → 実装 → 静的解析ループ(自動修正・最大3周)→ セルフレビュー → 完了報告 まで自律実行。
フロー定義と停止条件は .github/skills/autonomous-dev/SKILL.md に一元化。

### B. 手動モード(段階ごとに人が確認したい場合)
1. `/spec <要求>` → 仕様書起案 → 人が承認
2. 実装依頼(vba-developer / csharp-developer)
3. `/review` → 指摘対応 → コミット

### モデル選択方針(コスト最適化)
| フェーズ | 階層 | Claude Code | Copilot |
|---|---|---|---|
| 要求分析・仕様起案・レビュー | 高性能 | opus(agents の model で自動) | モデルピッカーで選択 |
| 実装 | 中性能 | sonnet(同上) | 同上 |
| 静的解析指摘の定型修正 | 低性能 | haiku(static-fixer) | 同上 |

Claude Code は .claude/agents/*.md の `model:` フロントマターで自動適用。
Copilot の .agent.md にも `# model:` 行を用意済み(コメントアウト)。環境のモデル一覧の正式名称に書き換えて有効化する(名称不一致だとエージェントが読み込めない環境があるため既定は無効)。

## トークン節約の仕組み
- 常駐するのは AGENTS.md(短文)とスキルの説明文のみ。
- 規約本体(SKILL.md)は関連タスク時のみロード。
- 指示の重複を排し、参照(リンク)で一元管理。

## 剪定運用(棚卸・自動化済み)
設定ファイルは毎リクエストでコンテキストに乗り続けるコスト。棚卸は2段構え:
1. 機械監査(全自動): 毎月1日に .github/workflows/instruction-audit.yml が scripts/audit_instructions.py を実行し、壊れた参照・行数予算超過・重複行・.claude⇔.github ペア不整合・フロントマター欠落を検出して Issue を起票。手動実行: `python scripts/audit_instructions.py --scan .`
2. 意味監査(半自動): Issue が立ったら(または月1回)`/audit-instructions` を実行。instruction-auditor エージェントが「古い仕様の残骸・重複・矛盾・肥大化・実効性・なぜの欠如」を横断監査し、承認後に修正まで実施。報告書は docs/audit/ に蓄積。

判断基準は .github/skills/instruction-audit/SKILL.md に一元化:
- AIが繰り返し無視したルール → 表現を強める or スキル/Hooksへ移す
- AGENTS.md が40行超 → スキルへ分離
- 各行に「この行を消すとAIが誤動作するか?」テスト。Noなら削除
- 「なぜ」が不明なルール → 理由を1句添える(遵守率・応用力が向上)

## 静的解析の自動検査(Hooks + CI)
3層で規約を機械的に保証する(指示文より確実):
| 層 | 設定ファイル | 動作 |
|---|---|---|
| Claude Code | .claude/settings.json | 編集直後に検査。違反はAIへ差し戻し自動修正させる(exit 2) |
| Copilot (CLI/coding agent/VS Code) | .github/hooks/doxygen.json | 編集直後に検査結果を通知 |
| CI (最終ゲート) | .github/workflows/doxygen-check.yml | PR/push時に全ファイル検査。人間のコミットも対象 |

検査本体は2本(要 Python 3.8+)。全層で共用:
- scripts/check_doxygen.py: C#=public類に `///` ヘッダ必須 / VBA=Public プロシージャに `'!` ヘッダ+Option Explicit 必須
- scripts/lint_vba.py: エラー握りつぶし(On Error Resume Next 放置)・秘密情報ハードコード・暗黙Variant・Select/Activate依存・ScreenUpdating未復帰を検出
- 手動実行: `python scripts/check_doxygen.py --scan .` / `python scripts/lint_vba.py --scan .`
- Windows で python3 コマンドが無い場合は設定内の python3 を python に読み替え(Claude Code側はフォールバック記述済み)

## プロジェクト知識ファイル(docs/)
| ファイル | 内容 | 作成優先度 |
|---|---|---|
| docs/glossary.md | 用語⇔コード対応表。AIの誤解釈防止の要 | 1(最優先) |
| docs/schema.md | テーブル/シート構造+サンプル1行 | 2 |
| docs/business-rules.md | 計算・判定ルール。例外を優先記載 | 3 |
| docs/domain/ | 領域別詳細(_template.md を複製) | 必要時 |
| docs/knowledge/ | ハマりどころ(_template.md を複製) | 随時 |

運用: AIが用語を誤解したら会話で訂正するだけでなく glossary.md に1行追加(会話は消えるがファイルは残る)。
頻出領域は .github/skills/_domain-template/ を複製して領域スキル化すると、関連タスク時のみ自動ロードされる。
