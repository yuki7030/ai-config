#!/usr/bin/env python3
"""subagent_status.py — 実行中サブエージェントの生存監視CLI (v2).

v2 変更点(設計欠陥の修正):
transcript のレコードは「ターン完了時」に書き込まれるため、
モデルが長考生成している最中は transcript が更新されない。
つまり transcript の実測だけでは「長考中」と「停滞」を区別できない。
v1 の WAITING / STALLED? 分類は長考中(=一番知りたい瞬間)に
誤判定するため廃止し、実測から言えることだけを表示する:

  STARTING  transcript 未生成(起動直後)
  ACTIVE    直近90秒以内にターン完了イベントあり(ツール呼び出しが流れている)
  LONG-TURN 90秒以上イベントなし = 長考生成中 or 停滞(区別不能)。
            LAST-EVT の経過秒と ELAPSED を見て判断する

補助情報として直近完了ターンの出力レート(tok/s)を表示する。
これは「直前のターン」の事後値であり現在の状態の証明ではない。

usage:
  python3 subagent_status.py            # 1回表示
  python3 subagent_status.py --watch    # 2秒間隔で更新
  python3 subagent_status.py --json     # 機械可読
"""
import json
import sys
import time
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import subagent_common as sc

ACTIVE_WINDOW_SEC = 90.0


def parse_ts(ts) -> float | None:
    if not ts:
        return None
    try:
        return datetime.fromisoformat(
            str(ts).replace("Z", "+00:00")).timestamp()
    except ValueError:
        return None


def tail_metrics(transcript: str) -> dict:
    m = {"exists": False, "last_event_age": None, "out_tokens": 0,
         "last_rate": None, "last_tool": "", "events": 0}
    p = Path(transcript) if transcript else None
    if not p or not p.exists():
        return m
    m["exists"] = True
    m["last_event_age"] = time.time() - p.stat().st_mtime

    prev_ts = None
    last_gap = None
    last_out = 0
    try:
        with open(p, encoding="utf-8", errors="replace") as f:
            for line in f:
                try:
                    rec = json.loads(line)
                except json.JSONDecodeError:
                    continue  # 書き込み途中の行はスキップ
                if not isinstance(rec, dict):
                    continue
                m["events"] += 1
                t = parse_ts(rec.get("timestamp"))
                if rec.get("type") == "assistant":
                    msg = rec.get("message") or {}
                    usage = msg.get("usage") or {}
                    try:
                        out = int(usage.get("output_tokens") or 0)
                    except (TypeError, ValueError):
                        out = 0
                    m["out_tokens"] += out
                    if t is not None and prev_ts is not None:
                        last_gap = max(t - prev_ts, 0.001)
                        last_out = out
                    for c in (msg.get("content") or []):
                        if isinstance(c, dict) and c.get("type") == "tool_use":
                            m["last_tool"] = c.get("name", "")
                if t is not None:
                    prev_ts = t
    except OSError:
        return m
    if last_gap:
        m["last_rate"] = last_out / last_gap
    return m


def classify(met: dict, transcript: str) -> str:
    if not met["exists"]:
        return "STARTING"
    if sc.transcript_finished(transcript):
        return "DONE?"  # 完了済み(SubagentStop未発火)→ 刈り取り対象
    age = met["last_event_age"] or 0
    if age <= ACTIVE_WINDOW_SEC:
        return "ACTIVE"
    return "LONG-TURN"  # 長考中 or 停滞(transcriptからは区別不能)


def snapshot() -> list[dict]:
    rows = []
    now = time.time()
    for agent_id, entry in sc.running_set(now).items():
        met = tail_metrics(entry.get("transcript", ""))
        state = classify(met, entry.get("transcript", ""))
        if state == "DONE?":
            # 完了済みを検出したら合成 stop で閉じる(reap)。
            # 今回の表示には DONE? として残し、次回から消える
            sc.reap(agent_id, entry.get("agent_type", "?"))
        rows.append({
            "agent": entry.get("agent_type", "?"),
            "id": agent_id[:8],
            "elapsed_s": round(now - entry.get("started_at", now)),
            "state": state,
            "last_event_s": (round(met["last_event_age"])
                             if met["last_event_age"] is not None else None),
            "out_tokens": met["out_tokens"],
            "last_rate": (round(met["last_rate"], 1)
                          if met["last_rate"] is not None else None),
            "last_tool": met["last_tool"],
        })
    return rows


def render(rows: list[dict]) -> str:
    if not rows:
        return "(実行中のサブエージェントなし)"
    hdr = f"{'AGENT':<12} {'STATE':<10} {'ELAPSED':>8} {'LAST-EVT':>9} " \
          f"{'OUT-TOK':>8} {'PREV-RATE':>10} {'LAST-TOOL':<12} ID"
    lines = [hdr, "-" * len(hdr)]
    for r in rows:
        le = f"{r['last_event_s']}s" if r["last_event_s"] is not None else "-"
        rate = f"{r['last_rate']}t/s" if r["last_rate"] is not None else "-"
        lines.append(
            f"{r['agent']:<12} {r['state']:<10} {r['elapsed_s']:>7}s "
            f"{le:>9} {r['out_tokens']:>8} {rate:>10} "
            f"{r['last_tool']:<12} {r['id']}")
    lines.append("")
    lines.append("LONG-TURN = 長考生成中 or 停滞(transcriptでは区別不能)。")
    lines.append("ELAPSED が伸び続けて LAST-EVT だけ増える場合は長考の可能性大。")
    lines.append("DONE? = 完了済みだが SubagentStop 未発火(background既知問題)。"
                 "刈り取り済みで次回表示から消える。")
    return "\n".join(lines)


def main() -> int:
    args = sys.argv[1:]
    if "--json" in args:
        print(json.dumps(snapshot(), ensure_ascii=False, indent=1))
        return 0
    if "--watch" in args:
        try:
            while True:
                print("\x1b[2J\x1b[H" + render(snapshot()))
                time.sleep(2)
        except KeyboardInterrupt:
            return 0
    print(render(snapshot()))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except BrokenPipeError:
        sys.exit(0)
