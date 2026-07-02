#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VBA静的検査(Claude Code hook / Copilot hook / CI 共用)。

AGENTS.md の NEVER(エラー握りつぶし・秘密情報ハードコード)と
vba-coding スキルの規約を機械的に検査する。使い方は check_doxygen.py と同一:
  --hook claude  : Claude Code PostToolUse。違反時 exit 2 (stderrがAIへ差し戻される)
  --hook copilot : Copilot postToolUse。違反時 メッセージ出力 + exit 0 (CIが最終ゲート)
  --scan PATH... : CI/手動。指定パス以下を全検査。違反時 exit 1

検査内容 (.bas/.cls/.frm):
  E1: On Error Resume Next の放置(同一プロシージャ内で On Error GoTo により復帰しない)
  E2: 秘密情報らしき文字列のハードコード(password / pwd / secret / apikey / token)
  E3: 型宣言のない Dim(暗黙 Variant)
  W1: .Select / .Activate / Selection への依存(遅く壊れやすい)
  W2: ScreenUpdating / EnableEvents を False にしたまま True に戻す行がない
"""
import json
import re
import sys
from pathlib import Path

VBA_EXT = {".bas", ".cls", ".frm"}

PROC_START = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property)\s+\w+", re.I)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
RESUME_NEXT = re.compile(r"^\s*On\s+Error\s+Resume\s+Next\b", re.I)
ON_ERROR_GOTO = re.compile(r"^\s*On\s+Error\s+GoTo\b", re.I)
SECRET = re.compile(r'(?:\b(?:password|passwd|pwd|secret|apikey|api_key|token)\s*=\s*"[^"]+"'
                    r'|"[^"]*\b(?:password|pwd)\s*=\s*[^";][^"]*")', re.I)
DIM_LINE = re.compile(r"^\s*(?:Dim|Private|Public|Static)\s+(?!Sub|Function|Property|Const|Type|Enum|Declare|WithEvents)(.+)$", re.I)
SELECT_ACTIVATE = re.compile(r"(\.\s*(Select|Activate)\s*$|\bSelection\s*\.)", re.I)
GUARD_OFF = {
    "ScreenUpdating": (re.compile(r"\.\s*ScreenUpdating\s*=\s*False", re.I),
                       re.compile(r"\.\s*ScreenUpdating\s*=\s*True", re.I)),
    "EnableEvents": (re.compile(r"\.\s*EnableEvents\s*=\s*False", re.I),
                     re.compile(r"\.\s*EnableEvents\s*=\s*True", re.I)),
}


def strip_comment(line: str) -> str:
    """文字列リテラル内を除き、' 以降のコメントを取り除く。"""
    out = []
    in_str = False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            break
        out.append(ch)
    return "".join(out)


def _check_dim(code: str):
    """Dim行から型宣言のない変数名を返す(暗黙Variant検出)。"""
    m = DIM_LINE.match(code)
    if not m:
        return []
    body = re.sub(r"\([^)]*\)", "()", m.group(1))  # 配列添字内のカンマを無効化
    bad = []
    for part in body.split(","):
        part = part.strip()
        if part and not re.search(r"\bAs\s+\w", part, re.I):
            bad.append(part.split("(")[0].strip())
    return bad


def check_file(path: Path):
    """ファイルを検査し違反リスト[(行番号, メッセージ)]を返す。"""
    try:
        text = path.read_text(encoding="utf-8-sig", errors="replace")
    except OSError:
        return []
    lines = text.splitlines()
    issues = []
    resume_line = None  # 未復帰の On Error Resume Next の行番号

    for i, raw in enumerate(lines):
        code = strip_comment(raw)
        if not code.strip():
            continue
        if PROC_START.match(code):
            resume_line = None
        if RESUME_NEXT.match(code):
            resume_line = i + 1
        elif ON_ERROR_GOTO.match(code):
            resume_line = None
        if PROC_END.match(code) and resume_line is not None:
            issues.append((resume_line, "E1: On Error Resume Next が復帰されないまま放置されています(エラー握りつぶし)"))
            resume_line = None
        if SECRET.search(code):
            issues.append((i + 1, "E2: 秘密情報らしき文字列がハードコードされています(設定・資格情報ストアへ)"))
        for name in _check_dim(code):
            issues.append((i + 1, f"E3: {name}: 型宣言がありません(暗黙Variant)。As で型を明示"))
        if SELECT_ACTIVATE.search(code):
            issues.append((i + 1, "W1: Select/Activate/Selection への依存。オブジェクト直接参照に置換"))

    for prop, (off, on) in GUARD_OFF.items():
        for i, raw in enumerate(lines):
            if off.search(strip_comment(raw)):
                if not any(on.search(strip_comment(l)) for l in lines):
                    issues.append((i + 1, f"W2: {prop} = False が True に戻されていません"))
                break
    return sorted(issues)


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
            if node.lstrip().startswith("{"):
                try:
                    walk(json.loads(node))
                    return
                except json.JSONDecodeError:
                    pass
            p = Path(node)
            if p.suffix.lower() in VBA_EXT and p.is_file():
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
                f for f in p.rglob("*") if f.suffix.lower() in VBA_EXT
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

    body = "[VBA静的検査] 規約違反があります。修正してください:\n" + "\n".join(report)
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
