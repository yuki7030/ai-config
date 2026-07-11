#!/usr/bin/env python3
"""Stop hook: 探索委譲の検査 — read-only ツールが委譲なしに
連続 N 回続いた形跡があれば警告ファイルを書く(warn のみ、block しない).

Zenn記事版からの変更点:
- v2.1.63 で Task ツールが Agent に改名されたため両方を委譲成立と見なす
- 委譲成立ツールに SendMessage(subagent resume)も追加
- CLAUDE_DELEGATION_FEEDBACK=immediate で、警告を次回セッションではなく
  そのターン末尾の additionalContext として即時に Claude へ返せる。
  注意: 公式仕様では Stop の additionalContext は「会話を継続して
  Claude がフィードバックに対応できる」= 追加ターンが発生し得る。
  コスト増と引き換えの即時性なので既定は next_session のまま
"""
import json
import os
import re
import sys
from pathlib import Path

READONLY_TOOLS = {"Read", "Grep", "Glob", "Bash"}
EDIT_TOOLS = {"Edit", "Write", "MultiEdit", "NotebookEdit"}
DELEGATION_TOOLS = {"Agent", "Task", "TaskCreate", "SendMessage"}
FEEDBACK_MODE = os.environ.get(
    "CLAUDE_DELEGATION_FEEDBACK", "next_session").lower()
BASH_MUTATE_RE = re.compile(
    r"(?:^|[;&|]\s*)(?:rm|mv|cp|mkdir|touch|sed\s+-i"
    r"|git\s+(?:commit|push|add|reset|checkout))\b|>>?\s*\S"
)
WARN_FILE = Path.home() / ".claude" / "delegation_warn.md"


def scan_max_streak(tpath: str) -> int:
    max_streak = streak = 0
    with open(tpath, encoding="utf-8", errors="replace") as f:
        for line in f:
            if '"assistant"' not in line:
                continue
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            if rec.get("type") != "assistant":
                continue
            for c in rec.get("message", {}).get("content", []):
                if not isinstance(c, dict) or c.get("type") != "tool_use":
                    continue
                name, inp = c.get("name", ""), c.get("input") or {}
                if name in DELEGATION_TOOLS or name in EDIT_TOOLS:
                    streak = 0
                elif name in READONLY_TOOLS:
                    if name == "Bash" and BASH_MUTATE_RE.search(
                        str(inp.get("command") or "")
                    ):
                        streak = 0
                    else:
                        streak += 1
                        max_streak = max(max_streak, streak)
    return max_streak


def main() -> int:
    if os.environ.get("CLAUDE_DELEGATION_CHECK", "0") != "1":
        return 0
    data = json.load(sys.stdin)
    tpath = str(data.get("transcript_path") or "")
    if not os.path.exists(tpath):
        return 0
    threshold = int(os.environ.get("CLAUDE_DELEGATION_STREAK_N", "8"))
    if scan_max_streak(tpath) >= threshold:
        msg = ("探索の subagent 委譲忘れ: read-only ツールが委譲なしに"
               f"{threshold} 回以上連続。explorer / scanner への委譲を検討。")
        if FEEDBACK_MODE == "immediate":
            print(json.dumps({"hookSpecificOutput": {
                "hookEventName": "Stop",
                "additionalContext": msg,
            }}, ensure_ascii=False))
        else:
            WARN_FILE.write_text(msg + "\n", encoding="utf-8")
    elif WARN_FILE.exists():
        WARN_FILE.unlink()
    return 0


if __name__ == "__main__":
    sys.exit(main())
