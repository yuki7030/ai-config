#!/usr/bin/env python3
"""acceptance_check.py — 実環境の transcript に対するパーサ受け入れ検証.

合成データでの単体テストは「自分の想定に合わせた検証」に留まる
(循環性)。このスクリプトを実環境で実行し、実際の Claude Code の
JSONL スキーマに対してパーサが機能するかを実測で確認する。

検証内容:
- 親 transcript: 総行数 / JSON解釈可能率 / timestamp解釈可能率 /
  assistant レコードの usage.output_tokens 取得率
- subagent transcript(存在すれば): 同上 + tail_metrics が例外なく
  値を返すか

判定: 各取得率が90%未満なら NG(スキーマ不一致の可能性)。
subagent transcript が1つも見つからない場合は SKIP と報告する
(委任を1回起こしてから再実行)。

usage: python3 acceptance_check.py [--root ~/.claude/projects]
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import subagent_status as ss


def audit_file(p: Path) -> dict:
    """timestamp率は user/assistant レコード限定で測る.

    修正: summary/system 等の非対象レコードは timestamp を持たなくて
    正常であり、パーサも参照しない。全レコード分母だと実環境で
    偽陰性(84%NG等)を出す。
    """
    total = parsed = ua = with_ts = assistants = with_usage = 0
    try:
        fh = open(p, encoding="utf-8", errors="replace")
    except OSError:
        return {}
    with fh:
        for line in fh:
            if not line.strip():
                continue
            total += 1
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            if not isinstance(rec, dict):
                continue
            parsed += 1
            if rec.get("type") in ("user", "assistant"):
                ua += 1
                ts = rec.get("timestamp")
                if ts:
                    try:
                        datetime.fromisoformat(
                            str(ts).replace("Z", "+00:00"))
                        with_ts += 1
                    except ValueError:
                        pass
            if rec.get("type") == "assistant":
                assistants += 1
                usage = (rec.get("message") or {}).get("usage") or {}
                if usage.get("output_tokens") is not None:
                    with_usage += 1
    return {"total": total, "parsed": parsed, "ua": ua,
            "with_ts": with_ts,
            "assistants": assistants, "with_usage": with_usage}


def rate(a: int, b: int) -> float:
    return 100.0 * a / b if b else 0.0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=str(Path.home() / ".claude/projects"))
    a = ap.parse_args()
    root = Path(a.root).expanduser()

    parents = sorted(root.glob("*/*.jsonl"),
                     key=lambda p: p.stat().st_mtime, reverse=True)[:20]
    subs = sorted(root.glob("*/*/subagents/agent-*.jsonl"),
                  key=lambda p: p.stat().st_mtime, reverse=True)[:20]

    ok = True
    print("== 親 transcript ==")
    if not parents:
        print("[NG] 見つからない(root を確認)")
        return 1
    agg = {"total": 0, "parsed": 0, "ua": 0, "with_ts": 0,
           "assistants": 0, "with_usage": 0}
    for p in parents:
        r = audit_file(p)
        for k in agg:
            agg[k] += r.get(k, 0)
    pr, tr, ur = (rate(agg["parsed"], agg["total"]),
                  rate(agg["with_ts"], agg["ua"]),
                  rate(agg["with_usage"], agg["assistants"]))
    print(f"  {len(parents)}ファイル {agg['total']}行: "
          f"JSON {pr:.0f}% / timestamp(u/a限定) {tr:.0f}% "
          f"/ usage {ur:.0f}%")
    if min(pr, tr, ur) < 90:
        ok = False
        print("[NG] 取得率90%未満の項目あり → スキーマ不一致の可能性。"
              "measure_gaps / status の数値は信頼できない")
    else:
        print("[OK] 親 transcript のスキーマ前提は成立")

    print("== subagent transcript ==")
    if not subs:
        print("[SKIP] 未発見。委任を1回起こしてから再実行して確認すること")
    else:
        agg2 = {"total": 0, "parsed": 0, "ua": 0, "with_ts": 0,
                "assistants": 0, "with_usage": 0}
        for p in subs:
            r = audit_file(p)
            for k in agg2:
                agg2[k] += r.get(k, 0)
            try:
                ss.tail_metrics(str(p))  # 例外なく通ることの確認
            except Exception as e:  # noqa: BLE001
                ok = False
                print(f"[NG] tail_metrics 例外: {p.name}: {e}")
        pr2, tr2, ur2 = (rate(agg2["parsed"], agg2["total"]),
                         rate(agg2["with_ts"], agg2["ua"]),
                         rate(agg2["with_usage"], agg2["assistants"]))
        print(f"  {len(subs)}ファイル {agg2['total']}行: "
              f"JSON {pr2:.0f}% / timestamp {tr2:.0f}% / usage {ur2:.0f}%")
        if min(pr2, tr2, ur2) < 90:
            ok = False
            print("[NG] subagent transcript のスキーマ前提が不成立")
        else:
            print("[OK] subagent transcript のスキーマ前提は成立")

    print("\n判定:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
