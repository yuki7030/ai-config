#!/usr/bin/env python3
"""Notification hook: 許可待ち・入力待ちを OS トーストに転送する.

参考: note.com/kirimajiro/n/n778f03813945(Notification/Stopフック構成)
      qiita.com/PaoJaPao/items/ea62930ea13b8912e118(permission_prompt通知)

「ツール許可待ちで気づかず放置」は subagent 無音と同種の可視性問題。
Notification イベントの JSON({title, message, ...})をそのまま
トーストへ転送し、どのプロジェクトかを cwd の basename で示す。
matcher(permission_prompt / idle_prompt)は settings 側で絞る。
常に exit 0。
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import subagent_common as sc


def main() -> int:
    # kind は settings 側の matcher に対応して argv で受ける
    # (permission / idle)。省略時は kind なし = 汎用アイコン
    kind = sys.argv[1] if len(sys.argv) > 1 else ""
    try:
        data = json.load(sys.stdin)
    except json.JSONDecodeError:
        return 0
    title = str(data.get("title") or "Claude Code")
    body = str(data.get("message") or "")
    cwd = str(data.get("cwd") or "")
    if cwd:
        body = (body + f"  [{Path(cwd).name}]").strip()
    sc.fire_native_notify(title, body or "\u901a\u77e5", kind=kind)
    return 0


if __name__ == "__main__":
    sys.exit(main())
