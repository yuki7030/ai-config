#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AI指示ファイルの機械監査(棚卸)スクリプト。

意味的監査(旧仕様の残骸・矛盾・実効性)は instruction-auditor エージェントが担当。
本スクリプトは機械検出可能な問題のみを扱う。
  ERROR: 壊れたパス参照 / .claude ⇔ .github ペア不整合 / フロントマター欠落
  WARN : 行数予算超過 / ファイル間の重複行
  INFO : プレースホルダ残存 / 180日以上未更新
終了コード: ERROR/WARN があれば 1(CI検知用)。INFOのみ・指摘なしは 0。

使い方: python scripts/audit_instructions.py --scan . [--format github]
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
import time
from collections import defaultdict
from pathlib import Path

TARGETS = [
    "AGENTS.md",
    "CLAUDE.md",
    ".github/copilot-instructions.md",
    ".github/instructions/*.instructions.md",
    ".github/skills/*/SKILL.md",
    ".github/agents/*.agent.md",
    ".github/prompts/*.prompt.md",
    ".claude/agents/*.md",
    ".claude/commands/*.md",
]
BUDGETS = [  # (ファイル名サフィックス, 最大行数)
    # 常駐すべき行動ガードレール(停止条件・進捗報告の裏付け・上流判断・
    # 思考の質)を含むため90行を上限とする。参照的な運用詳細はスキルへ分離済み。
    ("AGENTS.md", 90),
    ("CLAUDE.md", 10),
    ("copilot-instructions.md", 10),
    ("SKILL.md", 100),
    (".instructions.md", 20),
    (".agent.md", 15),
    (".prompt.md", 15),
]
# セットアップ時に生成される等の理由で、リポジトリに存在しなくてよい参照
REF_ALLOW = {".claude/skills"}
# Claude Code 専用のオーケストレーション用エージェント(組み込みExplore上書き・
# サブエージェント委譲機構)。Copilot は同等の自動委譲を持たないため、
# .github/agents 側のペアを意図的に持たない。ペア整合チェックから除外する。
CLAUDE_ONLY_AGENTS = {"Explore", "explorer", "scanner", "reviewer", "challenger"}
PATH_RE = re.compile(r"(?<![\w/@])((?:docs|scripts|\.github|\.claude)/[\w\-./]+)")
PLACEHOLDER_RE = re.compile("<[^<>]{1,40}>|\uFF08例\uFF09")
STALE_DAYS = 180
DUP_MIN_LEN = 20
# 重複行チェックの対象(agents/commands/prompts はペアで相似となる設計のため除外)
DUP_FILES = ("AGENTS.md", "CLAUDE.md", "copilot-instructions.md",
             ".instructions.md", "SKILL.md")


def is_template(p: Path) -> bool:
    return "_template" in str(p) or "_domain-template" in str(p)


def frontmatter_keys(lines: list[str]) -> set[str] | None:
    if not lines or lines[0].strip() != "---":
        return None
    keys: set[str] = set()
    for line in lines[1:]:
        if line.strip() == "---":
            return keys
        m = re.match(r"^([A-Za-z][\w-]*):", line)
        if m:
            keys.add(m.group(1))
    return None


def required_keys(p: Path) -> set[str]:
    s = str(p).replace("\\", "/")
    if "/agents/" in s:
        return {"name", "description", "tools"}
    if ".instructions.md" in s:
        return {"applyTo"}
    if ".prompt.md" in s or "/commands/" in s:
        return {"description"}
    if s.endswith("SKILL.md"):
        return {"name", "description"}
    return set()


def last_commit_age_days(root: Path, p: Path) -> float | None:
    try:
        out = subprocess.run(
            ["git", "log", "-1", "--format=%ct", "--", str(p.relative_to(root))],
            cwd=root, capture_output=True, text=True, timeout=10,
        ).stdout.strip()
        return (time.time() - int(out)) / 86400 if out else None
    except Exception:
        return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--scan", default=".")
    ap.add_argument("--format", choices=["text", "github"], default="text")
    args = ap.parse_args()
    root = Path(args.scan).resolve()

    files: list[Path] = []
    for pat in TARGETS:
        files.extend(sorted(root.glob(pat)))
    findings: list[tuple[str, str, str]] = []  # (severity, file, message)

    def add(sev: str, p: Path | str, msg: str) -> None:
        rel = str(p.relative_to(root)) if isinstance(p, Path) else p
        findings.append((sev, rel.replace("\\", "/"), msg))

    dup_map: dict[str, set[str]] = defaultdict(set)

    for p in files:
        lines = p.read_text(encoding="utf-8", errors="replace").splitlines()

        # 1) 行数予算
        for suffix, budget in BUDGETS:
            if str(p).endswith(suffix):
                if len(lines) > budget:
                    add("WARN", p, f"行数超過: {len(lines)}行(予算{budget}行)。分離・削除を検討")
                break

        # 2) フロントマター必須キー
        req = required_keys(p)
        if req:
            keys = frontmatter_keys(lines)
            if keys is None:
                add("ERROR", p, "フロントマターが無い")
            else:
                for k in sorted(req - keys):
                    add("ERROR", p, f"フロントマターに {k} が無い")

        # 3) 壊れたパス参照
        in_fence = False
        for line in lines:
            if line.lstrip().startswith("```"):
                in_fence = not in_fence
                continue
            if in_fence:
                continue
            for m in PATH_RE.finditer(line):
                cand = m.group(1).rstrip("。、.,;:)」")
                if cand.endswith("-") or "YYYY" in cand or "XXX" in cand or re.search(r"/X\.", cand):
                    continue  # プレースホルダを含む参照
                if cand.rstrip("/") in REF_ALLOW:
                    continue
                if not (root / cand).exists():
                    add("ERROR", p, f"存在しないパスへの参照: {cand}")

        # 4) 重複行の収集
        if any(str(p).endswith(s) for s in DUP_FILES):
            in_fence = False
            for line in lines:
                if line.lstrip().startswith("```"):
                    in_fence = not in_fence
                    continue
                norm = re.sub(r"\s+", " ", line.strip().lstrip("- ").strip())
                if (in_fence or len(norm) < DUP_MIN_LEN
                        or norm.startswith(("#", "|", "<!--"))):
                    continue
                dup_map[norm].add(str(p.relative_to(root)).replace("\\", "/"))

        # 5) プレースホルダ残存(テンプレートは除外)
        if not is_template(p):
            hits, in_fence = 0, False
            for l in lines:
                if l.lstrip().startswith("```"):
                    in_fence = not in_fence
                    continue
                if in_fence or l.lstrip().startswith("<!--"):
                    continue
                if PLACEHOLDER_RE.search(re.sub(r"`[^`]*`", "", l)):
                    hits += 1
            if hits:
                add("INFO", p, f"未記入プレースホルダらしき記述 {hits} 箇所(\uFF08例\uFF09/<...>)")

        # 6) 長期未更新
        age = last_commit_age_days(root, p)
        if age and age > STALE_DAYS:
            add("INFO", p, f"{int(age)}日間未更新。内容の現行性を確認")

    # 7) .claude ⇔ .github ペア整合
    ca = {q.stem for q in (root / ".claude/agents").glob("*.md")}
    ga = {q.name[:-len(".agent.md")] for q in (root / ".github/agents").glob("*.agent.md")}
    for name in sorted(ca - ga - CLAUDE_ONLY_AGENTS):
        add("ERROR", f".claude/agents/{name}.md", "対応する .github/agents/*.agent.md が無い")
    for name in sorted(ga - ca):
        add("ERROR", f".github/agents/{name}.agent.md", "対応する .claude/agents/*.md が無い")
    cc = {q.stem for q in (root / ".claude/commands").glob("*.md")}
    gp = {q.name[:-len(".prompt.md")] for q in (root / ".github/prompts").glob("*.prompt.md")}
    for name in sorted(cc - gp):
        add("ERROR", f".claude/commands/{name}.md", "対応する .github/prompts/*.prompt.md が無い")
    for name in sorted(gp - cc):
        add("ERROR", f".github/prompts/{name}.prompt.md", "対応する .claude/commands/*.md が無い")

    for norm, owners in sorted(dup_map.items()):
        if len(owners) > 1:
            add("WARN", " / ".join(sorted(owners)), f"重複行: 「{norm[:60]}…」参照への置換を検討")

    # 出力
    order = {"ERROR": 0, "WARN": 1, "INFO": 2}
    findings.sort(key=lambda f: order[f[0]])
    n = {s: sum(1 for f in findings if f[0] == s) for s in order}
    if args.format == "github":
        print(f"## 指示ファイル機械監査({time.strftime('%Y-%m-%d')})")
        print(f"対象: {len(files)} ファイル / 指摘: {len(findings)} 件"
              f"(ERROR: {n['ERROR']} / WARN: {n['WARN']} / INFO: {n['INFO']})\n")
        if findings:
            print("| 重大度 | ファイル | 指摘 |")
            print("|---|---|---|")
            for sev, f, msg in findings:
                print(f"| {sev} | {f} | {msg} |")
    else:
        print(f"対象 {len(files)} ファイル / 指摘 {len(findings)} 件 "
              f"(ERROR: {n['ERROR']} / WARN: {n['WARN']} / INFO: {n['INFO']})")
        for sev, f, msg in findings:
            print(f"[{sev}] {f}: {msg}")
    return 1 if n["ERROR"] or n["WARN"] else 0


if __name__ == "__main__":
    sys.exit(main())
