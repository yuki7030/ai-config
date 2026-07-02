---
name: instruction-audit
description: AI指示ファイル(AGENTS.md/CLAUDE.md/copilot-instructions/agents/skills/prompts/commands)の監査・棚卸・整理・修正時に必ず使用。古い記述・重複・矛盾・肥大化の検出手順と修正基準を定義。
---

# 指示ファイル監査(棚卸)

## 対象
AGENTS.md / CLAUDE.md / .github/copilot-instructions.md / .github/instructions/ / .github/skills/ / .github/agents/ / .github/prompts/ / .claude/agents/ / .claude/commands/ / .claude/settings.json / docs/ のAI向け知識ファイル

## 手順
1. 機械監査: `python scripts/audit_instructions.py --scan .` を実行し結果を取り込む。
2. 全対象ファイルを読み、下記観点で横断調査する。
3. 監査報告を提示し、修正方針の承認を得てから修正する(誤参照・タイプミス等の軽微な修正は即時可)。
4. 報告書を docs/audit/YYYY-MM-DD.md に保存する。

## 監査観点(優先順)
1. 古い記述: 存在しないパス・コマンドへの参照。廃止済み機能・旧仕様の残骸。180日以上未更新のファイルは現行性を確認。
2. 重複: 同一ルールの複数ファイル記載。原則「AGENTS.md=常駐の最小共通則 / SKILL.md=詳細本体 / 他は参照のみ」。重複は参照に置換(ロード契機が異なる段階的要約は許容し、その旨を記録)。
3. 矛盾: ルール間の競合(優先順位未定義)。ペアファイル(.claude/agents/X.md ⇔ .github/agents/X.agent.md、.claude/commands ⇔ .github/prompts)の内容乖離。
4. 肥大化: AGENTS.md ≤40行 / SKILL.md ≤100行 / instructions ≤20行。モデルが確実に従える指示は150〜200個が上限。1行ごとに「この行を消すとAIが誤動作するか?」を自問し、Noなら削除。
5. 実効性: 「なぜ」の無い規則には理由を1句添える。曖昧語(「適切に」等)を具体化。AIが繰り返し無視するルールは Hooks/CI への機械化を提案。
6. フロントマター: agents の tools 明示(最小権限)。description は自動委譲のトリガーになるため対象・拡張子・用途を具体的に。instructions の applyTo の妥当性。
7. 未記入プレースホルダ: テンプレート以外に残る `<...>`・(例)・空セクション。

## 修正基準
- 削除 > 短縮 > 分離 > 追記 の優先順。追記は最終手段。
- ペア(.claude ⇔ .github)は必ず両方を同時に修正する。
- 仕様の正は docs/(glossary/schema/business-rules)。指示ファイルと食い違えば docs を正とする。

## 報告書式
```
# 指示ファイル監査 YYYY-MM-DD
対象: N ファイル / 指摘: N 件(High: n / Mid: n / Low: n / Info: n)
| # | 重大度 | ファイル | 指摘 | 対処 |
|---|---|---|---|---|
総評: <2行以内>
```
