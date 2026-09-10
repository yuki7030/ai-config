#!/usr/bin/env python3
"""PostToolUse/PostToolUseFailure フック: Bash の実行内容と出力をログへ追記する.

Claude が Bash ツールで叩いたコマンドの出力は Claude のツール結果として
返るだけで、ユーザーの統合ターミナルには一切流れない。ユーザーが
`tail -f` で追える単一のログへ集約することで、その断絶を埋める。

出力先: .claude/logs/claude-bash.log(gitignore 対象)
常に exit 0(ログ取得の失敗でツール実行を止めないため)。
"""
import json
import os
import sys
from datetime import datetime
from pathlib import Path

# 1回の実行あたりの記録上限。ビルドログ等でログが数MB級に膨らむと
# tail -f 側が読みづらくなるため、頭から切り詰める
MAX_CHARS = 4000

LOG_RELATIVE = Path(".claude") / "logs" / "claude-bash.log"


def extract_output(response: object) -> str:
    """tool_response から標準出力・標準エラー相当のテキストを取り出す.

    Bash ツールの応答形状はバージョンで揺れるため、辞書なら既知キーを
    順に拾い、文字列ならそのまま使う。
    """
    if isinstance(response, str):
        return response
    if not isinstance(response, dict):
        return ""

    parts = []
    for key in ("stdout", "stderr", "output", "content", "result", "error"):
        value = response.get(key)
        if isinstance(value, str) and value.strip():
            parts.append(value.rstrip())
    if parts:
        return "\n".join(parts)
    return json.dumps(response, ensure_ascii=False)


def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def truncate(text: str) -> str:
    if len(text) <= MAX_CHARS:
        return text
    return text[:MAX_CHARS] + f"\n… (以降 {len(text) - MAX_CHARS} 文字を省略)"


def main() -> int:
    # sys.stdin は Windows ではロケール既定(CP932)でデコードされるため、
    # UTF-8 のフック入力に含まれる日本語が化ける。バイト列から明示的に読む
    try:
        raw = sys.stdin.buffer.read()
    except OSError:
        return 0
    try:
        data = json.loads(raw.decode("utf-8", errors="replace"))
    except (json.JSONDecodeError, ValueError):
        return 0

    if data.get("tool_name") != "Bash":
        return 0

    tool_input = data.get("tool_input") or {}
    command = normalize_newlines(str(tool_input.get("command") or "")).strip()
    if not command:
        return 0

    output = truncate(normalize_newlines(extract_output(data.get("tool_response"))).strip())
    stamp = datetime.now().strftime("%H:%M:%S")
    # PostToolUseFailure から呼ばれた場合を argv で区別する
    marker = "FAILED" if (len(sys.argv) > 1 and sys.argv[1] == "--failure") else "ok"

    cwd = data.get("cwd") or os.getcwd()
    log_path = Path(cwd) / LOG_RELATIVE
    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        # newline="" でテキストモードの改行変換を切る。既に CRLF を含む出力を
        # そのまま書くと \r\r\n になり、追尾時に1行おきの空行として見える
        with log_path.open("a", encoding="utf-8", errors="replace", newline="") as fp:
            fp.write(f"\n[{stamp}] $ {command}\n")
            if output:
                fp.write(output + "\n")
            fp.write(f"--- {marker}\n")
    except OSError:
        return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())
