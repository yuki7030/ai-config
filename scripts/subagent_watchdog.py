#!/usr/bin/env python3
"""subagent_watchdog.py — 途中経過通知 + 完了刈り取り(reaper).

v2.3: background subagent では SubagentStop が発火しない
(実測 + Issue #33049/#27755/#58637)ため、watchdog に reaper を統合。
各サイクルで transcript の完了判定を行い、完了していれば
合成 stop を追記 → 完了通知 → 終了する。

usage: subagent_watchdog.py <agent_id> <interval_sec> [--once]
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import subagent_common as sc

import os
MAX_PINGS = 5
NOTIFY_SEC = int(os.environ.get("CLAUDE_SUBAGENT_NOTIFY_SEC", "30"))
LOG_FILE = Path.home() / ".claude" / "subagent_activity.log"


def log(line: str) -> None:
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {line}\n")
    except OSError:
        pass


def main() -> int:
    if len(sys.argv) < 3:
        return 0
    agent_id = sys.argv[1]
    interval = max(int(sys.argv[2]), 1)
    once = "--once" in sys.argv

    for _ in range(MAX_PINGS):
        time.sleep(interval)
        entry = sc.running_set().get(agent_id)
        if entry is None:
            return 0  # 正規の SubagentStop 済み → 何もしない

        # 完了刈り取り: transcript が最終回答で終わっていれば閉じる
        if sc.transcript_finished(entry.get("transcript", "")):
            try:
                mtime = Path(entry["transcript"]).stat().st_mtime
                elapsed = max(mtime - entry["started_at"], 0)
            except OSError:
                elapsed = time.time() - entry["started_at"]
            sc.reap(agent_id, entry["agent_type"])
            log(f"REAP  {entry['agent_type']} id={agent_id} "
                f"elapsed={elapsed:.0f}s (SubagentStop未発火を刈り取り)")
            proj = Path(entry.get("cwd", "")).name if entry.get("cwd") \
                else ""
            suffix = f"  [{proj}]" if proj else ""
            if elapsed >= NOTIFY_SEC:
                sc.fire_native_notify(
                    "Claude Code",
                    f"subagent {entry['agent_type']} "
                    f"\u5b8c\u4e86 ({elapsed:.0f}s){suffix}", kind="reaped")
            return 0

        elapsed = time.time() - entry["started_at"]
        proj = Path(entry.get("cwd", "")).name if entry.get("cwd") else ""
        suffix = f"  [{proj}]" if proj else ""
        sc.fire_native_notify(
            "Claude Code",
            f"subagent {entry['agent_type']} \u307e\u3060\u5b9f\u884c\u4e2d "
            f"({elapsed/60:.1f}\u5206\u7d4c\u904e){suffix}", kind="progress")
        if once:
            return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
