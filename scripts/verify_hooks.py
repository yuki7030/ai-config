#!/usr/bin/env python3
"""verify_hooks.py — VS Code 拡張でフックが発火しているかの検証.

背景: 拡張内で settings.json の hooks が発火しないという報告がある
(anthropics/claude-code Issue #21736)。環境・バージョン依存のため、
導入後にこのカナリアで実測確認する。

手順:
 1. VS Code の Claude Code パネルで
    「explorer subagent で README.md を1ファイルだけ調査して」等、
    委任が起きる軽い依頼を投げる
 2. 完了後にこのスクリプトを実行:
    python3 ~/.claude/scripts/verify_hooks.py
"""
import json
import sys
import time
from pathlib import Path

LOG = Path.home() / ".claude" / "subagent_activity.log"
EVENTS = Path.home() / ".claude" / "subagent_events.jsonl"
RECENT_SEC = 15 * 60


def main() -> int:
    print("== hooks 発火検証 ==")
    ok = True

    if LOG.exists():
        age = time.time() - LOG.stat().st_mtime
        lines = LOG.read_text(encoding="utf-8").strip().splitlines()
        print(f"[OK] activity log あり(最終更新 {age/60:.1f} 分前, "
              f"{len(lines)} 行)")
        for ln in lines[-5:]:
            print(f"     {ln}")
        if age > RECENT_SEC:
            print("[!] 直近15分の記録なし。委任直後に再実行して確認を")
    else:
        ok = False
        print("[NG] activity log が存在しない → SubagentStart/Stop フックが"
              "発火していない可能性")

    if EVENTS.exists():
        n = sum(1 for _ in open(EVENTS, encoding="utf-8"))
        print(f"[OK] events log あり({n} イベント)")
    else:
        print("[--] events log なし(一度も委任が起きていなければ正常)")

    if not ok:
        print()
        print("対処:")
        print(" 1. パネルで /hooks を実行し、SubagentStart/SubagentStop が"
              "一覧に出るか確認(出ない→settings.json のマージ漏れ)")
        print(" 2. 一覧に出るのに発火しない→拡張の既知問題の可能性。")
        print("    回避策: 拡張設定の Use Terminal を有効化するか、")
        print("    統合ターミナルで claude を起動(CLI では発火が確認済み)")
        print(" 3. claude --debug でフック実行ログを確認")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
