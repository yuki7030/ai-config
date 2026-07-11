#!/usr/bin/env python3
"""measure_gaps.py — 停滞 gap 分布の実測(導入前後の効果測定).

Zenn 記事の集計手法をツール化したもの。ハーネス導入の効果は
このスクリプトで before/after を比較して判断する。

出力: モデル別に 件数 / 中央値 / p90 / p99 / 最大 / >120s率 と、
60秒超 gap の 待ち・生成中 内訳(出力レート 5 tok/s 閾値)。

usage:
  python3 measure_gaps.py [--days 7] [--root ~/.claude/projects] [--json]
"""
import argparse
import json
import time
from datetime import datetime
from pathlib import Path


def parse_ts(ts) -> float | None:
    try:
        return datetime.fromisoformat(
            str(ts).replace("Z", "+00:00")).timestamp()
    except (ValueError, TypeError):
        return None


def collect(root: Path, since: float) -> dict:
    gaps: dict[str, list[tuple[float, int]]] = {}  # model -> [(gap, out_tok)]
    for f in root.glob("**/*.jsonl"):
        try:
            if f.stat().st_mtime < since:
                continue
        except OSError:
            continue
        prev_ts = None
        try:
            fh = open(f, encoding="utf-8", errors="replace")
        except OSError:
            continue
        with fh:
            for line in fh:
                try:
                    rec = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if not isinstance(rec, dict):
                    continue
                t = parse_ts(rec.get("timestamp"))
                if t is None:
                    continue
                if rec.get("type") == "user":
                    prev_ts = t
                elif rec.get("type") == "assistant" and prev_ts is not None:
                    msg = rec.get("message") or {}
                    gap = t - prev_ts
                    if 0 <= gap < 3600:
                        usage = msg.get("usage") or {}
                        try:
                            out = int(usage.get("output_tokens") or 0)
                        except (TypeError, ValueError):
                            out = 0
                        gaps.setdefault(
                            str(msg.get("model", "?")), []).append((gap, out))
                    prev_ts = None
    return gaps


def pct(sorted_vals: list[float], p: float) -> float:
    if not sorted_vals:
        return 0.0
    i = min(int(len(sorted_vals) * p), len(sorted_vals) - 1)
    return sorted_vals[i]


def summarize(gaps: dict) -> list[dict]:
    rows = []
    for model, pairs in sorted(gaps.items()):
        vals = sorted(g for g, _ in pairs)
        n = len(vals)
        over120 = sum(1 for g in vals if g > 120)
        long_pairs = [(g, o) for g, o in pairs if g > 60]
        waiting = sum(1 for g, o in long_pairs if o / g < 5.0)
        rows.append({
            "model": model, "n": n,
            "p50": round(pct(vals, 0.50), 1),
            "p90": round(pct(vals, 0.90), 1),
            "p99": round(pct(vals, 0.99), 1),
            "max": round(vals[-1], 1) if vals else 0,
            "over120s_pct": round(100 * over120 / n, 2) if n else 0,
            "gap60s_total": len(long_pairs),
            "gap60s_waiting": waiting,
            "gap60s_generating": len(long_pairs) - waiting,
        })
    return rows


def render(rows: list[dict]) -> str:
    if not rows:
        return "(対象期間に計測対象のレコードなし)"
    hdr = f"{'MODEL':<22} {'N':>6} {'P50':>6} {'P90':>7} {'P99':>7} " \
          f"{'MAX':>7} {'>120s':>6} {'60s+ 待ち/生成中':>16}"
    lines = [hdr, "-" * len(hdr)]
    for r in rows:
        lines.append(
            f"{r['model']:<22} {r['n']:>6} {r['p50']:>5}s {r['p90']:>6}s "
            f"{r['p99']:>6}s {r['max']:>6}s {r['over120s_pct']:>5}% "
            f"{r['gap60s_waiting']:>7}/{r['gap60s_generating']}")
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=7)
    ap.add_argument("--root", default=str(Path.home() / ".claude/projects"))
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    rows = summarize(collect(Path(a.root).expanduser(),
                             time.time() - a.days * 86400))
    if a.json:
        print(json.dumps(rows, ensure_ascii=False, indent=1))
    else:
        print(render(rows))
    return 0


if __name__ == "__main__":
    main()
