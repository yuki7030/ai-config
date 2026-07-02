#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Doxygenヘッダコメント検査(Claude Code hook / Copilot hook / CI 共用)。

使い方:
  --hook claude  : Claude Code PostToolUse。違反時 exit 2 (stderrがAIへ差し戻される)
  --hook copilot : Copilot postToolUse。違反時 メッセージ出力 + exit 0 (CIが最終ゲート)
  --scan PATH... : CI/手動。指定パス以下を全検査。違反時 exit 1

検査内容:
  C#  (.cs)            : public/protected/internal のクラス・メソッド等に /// ヘッダ必須
  VBA (.bas/.cls/.frm) : Public Sub/Function/Property に '! ヘッダ必須、Option Explicit 必須
"""
import json
import re
import sys
from pathlib import Path

CS_EXT = {".cs"}
VBA_EXT = {".bas", ".cls", ".frm"}
TARGET_EXT = CS_EXT | VBA_EXT

# C#: publicなクラス系・メソッド系宣言(フィールド・自動プロパティは対象外にして誤検知を回避)
CS_DECL = re.compile(
    r"^\s*(?:\[[^\]]*\]\s*)*(public|protected|internal)\b"
    r"(?=.*(?:\(|\bclass\b|\binterface\b|\bstruct\b|\brecord\b|\benum\b|\bdelegate\b))"
)
CS_SKIP = re.compile(r"^\s*(//|\*|/\*)|=>\s*$")  # コメント行等
VBA_DECL = re.compile(r"^\s*Public\s+(Sub|Function|Property)\s+(\w+)", re.IGNORECASE)


def _prev_has_doc(lines, idx, marker):
    """宣言行の直前(空行・C#属性を飛ばす)にdocコメントがあるか。"""
    i = idx - 1
    while i >= 0:
        s = lines[i].strip()
        if s == "" or (marker == "///" and s.startswith("[")):
            i -= 1
            continue
        return s.startswith(marker)
    return False


def check_file(path: Path):
    """ファイルを検査し違反リスト[(行番号, メッセージ)]を返す。"""
    try:
        text = path.read_text(encoding="utf-8-sig", errors="replace")
    except OSError:
        return []
    lines = text.splitlines()
    issues = []
    ext = path.suffix.lower()

    if ext in CS_EXT:
        for i, line in enumerate(lines):
            if CS_SKIP.match(line) or not CS_DECL.match(line):
                continue
            if not _prev_has_doc(lines, i, "///"):
                issues.append((i + 1, "Doxygenヘッダ(/// @brief 等)がありません"))
    elif ext in VBA_EXT:
        if not re.search(r"^\s*Option\s+Explicit", text, re.IGNORECASE | re.MULTILINE):
            issues.append((1, "Option Explicit がありません"))
        for i, line in enumerate(lines):
            m = VBA_DECL.match(line)
            if m and not _prev_has_doc(lines, i, "'!"):
                issues.append((i + 1, f"{m.group(2)}: Doxygenヘッダ('! @brief 等)がありません"))
    return issues


def _paths_from_stdin_json():
    """hookのstdin JSONから、実在する対象拡張子のファイルパスを収集(形式差異に耐える)。"""
    try:
        data = json.load(sys.stdin)
    except (json.JSONDecodeError, OSError):
        return []
    found = []

    def walk(node):
        if isinstance(node, dict):
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)
        elif isinstance(node, str):
            # toolArgsがJSON文字列のことがあるため二重パースを試みる
            if node.lstrip().startswith("{"):
                try:
                    walk(json.loads(node))
                    return
                except json.JSONDecodeError:
                    pass
            p = Path(node)
            if p.suffix.lower() in TARGET_EXT and p.is_file():
                found.append(p)

    walk(data)
    return list(dict.fromkeys(found))


def main():
    args = sys.argv[1:]
    mode = None
    targets = []
    if args[:1] == ["--hook"] and len(args) >= 2:
        mode = args[1]
        targets = _paths_from_stdin_json()
    elif args[:1] == ["--scan"]:
        mode = "scan"
        for a in args[1:] or ["."]:
            p = Path(a)
            targets += [p] if p.is_file() else [
                f for f in p.rglob("*") if f.suffix.lower() in TARGET_EXT
            ]
    else:
        print(__doc__)
        return 0

    report = []
    for f in targets:
        for lineno, msg in check_file(f):
            report.append(f"{f}:{lineno}: {msg}")

    if not report:
        return 0

    body = "[Doxygen検査] 規約違反があります。修正してください:\n" + "\n".join(report)
    if mode == "claude":
        print(body, file=sys.stderr)
        return 2  # Claude Codeにブロック+フィードバック
    if mode == "copilot":
        print(body)  # Copilotセッションログへ(最終ゲートはCI)
        return 0
    print(body, file=sys.stderr)  # scan(CI)
    return 1


if __name__ == "__main__":
    sys.exit(main())
