---
name: agent-workflow
description: サブエージェントへの探索委譲・モデルルーティング(explorer/scanner/reviewer/Explore)と教訓の記録運用。読み取り専用の探索を委譲する時、マイルストーン検証時、修復・仕様矛盾が起きた時に参照。主に Claude Code 向け。
---

# サブエージェント運用ガイド

explorer / scanner / reviewer / challenger / Explore は **Claude Code 専用**の
サブエージェント委譲機構(`.claude/agents/`)。Copilot は同等の自動委譲を
持たないため、対応する `.github/agents/` を意図的に置かない。

## モデルルーティング

| タスク種別 | 実行先 | モデル/effort | 根拠 |
|---|---|---|---|
| 実装・設計・複雑なデバッグ | 親セッション | Opus (session) | 初回正確性が必要 |
| 意味理解を伴う探索(仕様調査/複数ファイル調査/Webリサーチ) | explorer | Sonnet / medium | 探索に長考は不要 |
| 機械的 grep・件数集計 | scanner | Haiku / low | 解釈不要 |
| マイルストーン後の独立検証 | reviewer | Sonnet / high | 新規コンテキストの検証は自己批評より優れる |
| コードベース探索(組み込みExplore) | Explore(上書き定義) | Sonnet / medium | v2.1.198以降Opus継承を防ぐ |

## 探索の委譲

- 読み取り専用の探索(複数ファイル grep / 仕様調査 /
  ログ調査 / Web リサーチ)は subagent へ委譲する
  - 意味理解を伴う調査 → explorer
  - 機械的な grep 列挙・件数集計 → scanner
- 読み取りツールを連続 8 回以上使う見込み、
  または 5 ファイル以上を読む探索は委譲必須
- 委譲したら subagent の完了を待たずに、独立して進められる作業を続ける。
  subagent が本題から逸れた場合のみ介入する
- subagent は構造化サマリ(path:line + 判定)のみ返す
- 編集・commit は親セッションが行う(subagent 内での mutation 禁止)
- 実装マイルストーンの完了ごとに reviewer subagent で仕様との差分を検証する

## 教訓の記録と読み出し

- 修復・仕様矛盾・reviewer の Critical 指摘が発生した場合のみ、
  解決後に docs/ai-lessons.md へ教訓を1〜3行で追記する
  (何が起きたか / 根本原因 / 次回の回避策)
- 同種のタスクに着手する前に docs/ai-lessons.md を読み、
  該当する教訓があれば計画に反映する
- 問題が発生しなかったタスクでは記録しない(記録自体のコスト抑制)
