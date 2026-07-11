#!/usr/bin/env python3
"""SubagentStart / SubagentStop hook: サブエージェント委任の可視化 (v2).

v2 変更点:
- state file → 追記専用イベントログ(並列フック実行での lost update 解消)
- systemMessage も通知閾値でゲート(並列委任時のスパム防止)
- 長時間実行 watchdog を Start 時に起動(途中経過の proactive 通知)
- Windows トーストは AUMID フォールバック付き(subagent_common)

通知経路 CLAUDE_SUBAGENT_NOTIFY: native(既定, VS Code拡張向け)/ osc / both / off
hook としては常に exit 0(委任を絶対に妨げない)。
"""
import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import subagent_common as sc

LOG_FILE = Path.home() / ".claude" / "subagent_activity.log"
NOTIFY_SEC = int(os.environ.get("CLAUDE_SUBAGENT_NOTIFY_SEC", "30"))
PROGRESS_SEC = int(os.environ.get("CLAUDE_SUBAGENT_PROGRESS_SEC", "120"))
ESC, BEL = "\x1b", "\x07"


def log(line: str) -> None:
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{ts} {line}\n")


def subagent_transcript_path(parent_transcript: str, agent_id: str) -> str:
    if not parent_transcript or not agent_id:
        return ""
    base = parent_transcript[:-6] if parent_transcript.endswith(".jsonl") \
        else parent_transcript
    return str(Path(base) / "subagents" / f"agent-{agent_id}.jsonl")


def title_sequence() -> str:
    running = [v["agent_type"] for v in sc.running_set().values()]
    label = ("CC \u25b6 " + ", ".join(sorted(running))) if running \
        else "CC \u2713 idle"
    return f"{ESC}]0;{label}{BEL}"


def notify_sequence(title: str, body: str) -> str:
    return f"{ESC}]9;{title}: {body}{BEL}"


def spawn_watchdog(agent_id: str) -> None:
    if PROGRESS_SEC <= 0:
        return
    sc.spawn_detached(  # Windows でも親フック終了に道連れにされない
        [sys.executable,
         str(Path(__file__).resolve().parent / "subagent_watchdog.py"),
         agent_id, str(PROGRESS_SEC)])


def main() -> int:
    try:
        data = json.load(sys.stdin)
    except json.JSONDecodeError:
        return 0

    event = data.get("hook_event_name", "")
    agent_id = str(data.get("agent_id") or "")
    agent_type = str(data.get("agent_type") or "unknown")
    out: dict = {}

    if event == "SubagentStart":
        sc.append_event(
            "start", agent_id, agent_type,
            session_id=str(data.get("session_id") or ""),
            transcript=subagent_transcript_path(
                str(data.get("transcript_path") or ""), agent_id))
        log(f"START {agent_type} id={agent_id}")
        spawn_watchdog(agent_id)
        if sc.NOTIFY_MODE in ("osc", "both"):
            out["terminalSequence"] = title_sequence()

    elif event == "SubagentStop":
        start = sc.find_start(agent_id) if agent_id else None
        already_reaped = bool(agent_id) and agent_id not in sc.running_set()
        sc.append_event("stop", agent_id, agent_type)
        elapsed = time.time() - start["ts"] if start else -1
        log(f"STOP  {agent_type} id={agent_id} elapsed={elapsed:.0f}s"
            + (" (reap済み)" if already_reaped else ""))
        long_run = elapsed >= NOTIFY_SEC and not already_reaped
        if sc.NOTIFY_MODE in ("osc", "both"):
            seq = title_sequence()
            if long_run:
                seq += notify_sequence(
                    "Claude Code",
                    f"{agent_type} \u5b8c\u4e86 ({elapsed:.0f}s)")
            out["terminalSequence"] = seq
        if long_run:
            if sc.NOTIFY_MODE in ("native", "both"):
                sc.fire_native_notify(
                    "Claude Code",
                    f"subagent {agent_type} \u5b8c\u4e86 ({elapsed:.0f}s)")
            # systemMessage は長時間実行のみ(スパム防止)
            out["systemMessage"] = (
                f"\u2713 subagent {agent_type} \u5b8c\u4e86 "
                f"({elapsed:.0f}s)")

    if out:
        print(json.dumps(out, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
