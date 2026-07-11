---
name: Explore
description: >-
  Fast read-only codebase exploration (overrides the built-in Explore
  to pin a lightweight model instead of inheriting the session model).
tools: Read, Grep, Glob, Bash
model: sonnet
effort: medium
color: cyan
background: true
---

# Explore (model-pinned override) — 導入は任意。トレードオフを理解して選ぶ

## 利益
v2.1.198 以降、組み込み Explore はメイン会話のモデルを継承する。
メインが Opus のセッションでは Explore も Opus で走り、
長考による無音・コスト増の一因になる。この定義で Sonnet に固定できる。

## コスト(公式ドキュメント明記の副作用)
組み込み Explore/Plan だけが CLAUDE.md と git status の読み込みを
スキップする最適化を持つ。同名カスタム定義で上書きすると:
- Explore 呼び出しごとに CLAUDE.md 全量 + git status が入力トークンに乗る
- 組み込みのチューニング済みプロンプト(thoroughness レベル対応)を失う

## 判断基準
- CLAUDE.md が小さい(数KB) & Opus セッションが多い → 導入する価値あり
- CLAUDE.md が大きい → 導入しない方が安い可能性。
  代わりに CLAUDE.md の委譲規則で explorer(本ハーネス)への
  委譲を促す方が副作用がない
- 導入する場合: このファイルを ~/.claude/agents/ にコピー

役割: 高速な読み取り専用のコードベース探索。
結果は構造化サマリ(path:line + 判定)で返し、raw dump は返さない。
