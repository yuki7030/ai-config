#!/usr/bin/env python3
"""subagent_common.py — ハーネス共通ロジック.

設計変更(v2): state file の load→modify→save をやめ、
追記専用イベントログ(JSONL)方式に変更した。

理由:
- hooks は並列実行される。並列サブエージェントの一斉ディスパッチ時、
  read-modify-write の state file は lost update を起こす。
  追記(1行の単発 write)は POSIX の O_APPEND で実用上アトミック
- SubagentStop が来ない異常終了(セッション kill 等)でも、
  「実行中」の再構成時に TTL で自然に除外できる(ゾンビ解消)

ログ肥大対策: 追記前にサイズ超過を検知したら os.replace で .1 に
ローテーション(アトミック)。再構成は .1 と現行の両方を読む。
"""
import json
import os
import platform
import subprocess
import time
from pathlib import Path

EVENTS_FILE = Path.home() / ".claude" / "subagent_events.jsonl"
ROTATE_BYTES = int(os.environ.get("CLAUDE_SUBAGENT_LOG_ROTATE", "524288"))
TTL_SEC = int(os.environ.get("CLAUDE_SUBAGENT_TTL_SEC", str(24 * 3600)))
NOTIFY_MODE = os.environ.get("CLAUDE_SUBAGENT_NOTIFY", "native").lower()
NOTIFY_FILE = os.environ.get("CC_NOTIFY_FILE", "")  # テスト用フック

# ---------- イベントログ ----------

def _unmatched_starts(events: list[dict]) -> list[dict]:
    """イベント列から未マッチ(実行中)の start を抽出する."""
    starts: dict[str, dict] = {}
    for e in events:
        eid = str(e.get("id") or "")
        if e.get("ev") == "start":
            starts[eid or f"noid-{e.get('ts')}"] = e
        elif e.get("ev") == "stop":
            if eid and eid in starts:
                starts.pop(eid)
            elif not eid:
                cand = [(v.get("ts", 0), k) for k, v in starts.items()
                        if v.get("agent") == e.get("agent")]
                if cand:
                    starts.pop(min(cand)[1])
    return list(starts.values())


def _read_file(p: Path) -> list[dict]:
    out = []
    try:
        with open(p, encoding="utf-8", errors="replace") as f:
            for line in f:
                try:
                    out.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    except OSError:
        pass
    return out


def _rotate() -> None:
    """サイズ超過時のローテーション.

    バグ修正(v2.1): 単純 rename だと未完了 start が .1 の上書きで消え、
    running_set の再構成が壊れる。rename 後、未マッチ start を
    新ファイルへ再追記して持ち越す(重複は再構成側で冪等)。
    """
    try:
        current = _read_file(EVENTS_FILE)
        carryover = _unmatched_starts(current)
        os.replace(EVENTS_FILE, EVENTS_FILE.with_suffix(".jsonl.1"))
        if carryover:
            with open(EVENTS_FILE, "a", encoding="utf-8") as f:
                for e in carryover:
                    f.write(json.dumps(e, ensure_ascii=False) + "\n")
    except OSError:
        pass


def append_event(ev: str, agent_id: str, agent_type: str,
                 session_id: str = "", transcript: str = "") -> None:
    EVENTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    try:
        if EVENTS_FILE.exists() and EVENTS_FILE.stat().st_size > ROTATE_BYTES:
            _rotate()
    except OSError:
        pass
    line = json.dumps({
        "ts": time.time(), "ev": ev, "id": agent_id,
        "agent": agent_type, "session": session_id,
        "transcript": transcript,
    }, ensure_ascii=False) + "\n"
    with open(EVENTS_FILE, "a", encoding="utf-8") as f:
        f.write(line)


def load_events() -> list[dict]:
    events = []
    for p in (EVENTS_FILE.with_suffix(".jsonl.1"), EVENTS_FILE):
        if not p.exists():
            continue
        try:
            with open(p, encoding="utf-8", errors="replace") as f:
                for line in f:
                    try:
                        events.append(json.loads(line))
                    except json.JSONDecodeError:
                        continue  # 競合等での破損行はスキップ
        except OSError:
            continue
    events.sort(key=lambda e: e.get("ts", 0))
    return events


def running_set(now: float | None = None) -> dict:
    """start − stop の差分から実行中集合を再構成する.

    - TTL(既定24h)超過の start はゾンビと見なし除外
    - id 欠落の stop は、同 agent_type の最古の未マッチ start に対応づける
    """
    now = now or time.time()
    running: dict[str, dict] = {}
    for e in load_events():
        eid = str(e.get("id") or "")
        if e.get("ev") == "start":
            key = eid or f"noid-{e.get('ts')}"
            running[key] = {
                "agent_type": e.get("agent", "?"),
                "started_at": e.get("ts", now),
                "session_id": e.get("session", ""),
                "transcript": e.get("transcript", ""),
            }
        elif e.get("ev") == "stop":
            if eid and eid in running:
                running.pop(eid)
            elif not eid:  # フォールバック: 同型の最古startを消す
                cand = [(v["started_at"], k) for k, v in running.items()
                        if v["agent_type"] == e.get("agent")]
                if cand:
                    running.pop(min(cand)[1])
    return {k: v for k, v in running.items()
            if now - v["started_at"] <= TTL_SEC}


def find_start(agent_id: str) -> dict | None:
    for e in load_events():
        if e.get("ev") == "start" and str(e.get("id")) == agent_id:
            return e
    return None


# ---------- 完了検知(フック非依存の reaper) ----------
# 背景: background subagent では SubagentStop が発火しない
# (実測 + Issue #33049 / #27755 / #58637)。フックに依存せず
# transcript の最終レコードから完了を判定する。
#
# ヒューリスティクス: subagent のループ構造上、tool_use を含まない
# assistant レコード(テキストのみ)は最終回答である。実行中の agent の
# 末尾は必ず tool_use 付き(次ツール待ち)か、まだ書かれていない。
# 誤読防止に「ファイルが grace 秒以上未更新」を併用する。

REAP_GRACE_SEC = float(os.environ.get("CLAUDE_SUBAGENT_REAP_GRACE", "10"))


def transcript_finished(transcript: str,
                        grace_sec: float | None = None) -> bool:
    grace = REAP_GRACE_SEC if grace_sec is None else grace_sec
    p = Path(transcript) if transcript else None
    if not p or not p.exists():
        return False
    try:
        if time.time() - p.stat().st_mtime < grace:
            return False  # 書き込み直後は判定しない
    except OSError:
        return False
    last_assistant = None
    try:
        with open(p, encoding="utf-8", errors="replace") as f:
            for line in f:
                try:
                    rec = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if isinstance(rec, dict) and rec.get("type") == "assistant":
                    last_assistant = rec
    except OSError:
        return False
    if last_assistant is None:
        return False
    content = (last_assistant.get("message") or {}).get("content") or []
    has_tool_use = any(isinstance(c, dict) and c.get("type") == "tool_use"
                       for c in content)
    return not has_tool_use  # テキストのみ = 最終回答済み


def reap(agent_id: str, agent_type: str) -> None:
    """完了と判定した agent に合成 stop を追記して閉じる."""
    append_event("stop", agent_id, agent_type)

# ---------- OS ネイティブ通知 ----------

# AUMID フォールバック付きトースト:
#  1. 登録済み AUMID(Windows PowerShell)を優先する。未登録 AUMID
#     ("Claude Code" 等)は CreateToastNotifier/Show が例外を出さずに
#     表示だけされない(サイレント失敗)。先頭に置くと後続に到達せず
#     何も表示されない — 2026-07-11 実測で検出・修正
#  2. 万一 PowerShell AUMID が無い環境向けに "Claude Code" を後段に残す
#  3. Show() はトーストを非同期にキューするだけ。detached プロセスが
#     直後に終了すると配信前に消えるため、末尾で待機して配信猶予を与える
WIN_TOAST_PS = r"""
$ErrorActionPreference = 'Stop'
# WinRT 型リテラルは1行で書く(カンマ後に改行すると PowerShell が
# 「型名にアセンブリ名が指定されていません」で構文エラーになり、
# トーストが一切表示されない — 2026-07-11 実測で検出・修正)
$null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime]
$null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType=WindowsRuntime]
$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml(@"
<toast><visual><binding template="ToastGeneric">
<text>$($env:CC_TOAST_TITLE)</text><text>$($env:CC_TOAST_BODY)</text>
</binding></visual></toast>
"@)
$toast = New-Object Windows.UI.Notifications.ToastNotification $xml
$ids = @(
  '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe',
  'Claude Code'
)
foreach ($id in $ids) {
  try {
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier(
      $id).Show($toast)
    break
  } catch { continue }
}
Start-Sleep -Seconds 5
"""


# ---------- プロセス分離起動 ----------

def _popen_kwargs(interactive: bool) -> dict:
    """分離起動の Popen kwargs.

    バグ修正: start_new_session は POSIX 専用で Windows では無視される。
    フックは timeout で数秒後に終了するため、Windows では
    watchdog(スリープ中)が生き残るための creationflags が要る。

    interactive で2系統に分ける(2026-07-11 実測で必要と判明):
    - False(watchdog 等・既定): DETACHED_PROCESS で確実に分離し 120 秒生存。
    - True(通知トースト): DETACHED_PROCESS だと非対話セッション扱いになり
      トーストが対話デスクトップに配信されず、例外も出ず無表示になる。
      CREATE_NO_WINDOW なら対話セッション(WinSta0)のまま・ウィンドウ非表示。
      Windows は親終了で子を道連れにしないので短時間の配信には十分。
    """
    kw: dict = {"stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL,
                "stdin": subprocess.DEVNULL, "close_fds": True}
    if platform.system() == "Windows":
        NEW_GROUP = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP",
                            0x00000200)
        if interactive:
            CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW",
                                       0x08000000)
            kw["creationflags"] = CREATE_NO_WINDOW | NEW_GROUP
        else:
            DETACHED_PROCESS = getattr(subprocess, "DETACHED_PROCESS",
                                       0x00000008)
            kw["creationflags"] = DETACHED_PROCESS | NEW_GROUP
    else:
        kw["start_new_session"] = True
    return kw


def detached_popen_kwargs() -> dict:
    return _popen_kwargs(interactive=False)


def spawn_detached(argv: list, extra_env: dict | None = None,
                   interactive: bool = False) -> bool:
    try:
        subprocess.Popen(argv, env={**os.environ, **(extra_env or {})},
                         **_popen_kwargs(interactive))
        return True
    except OSError:
        return False


def native_notify_cmd(title: str, body: str):
    sysname = platform.system()
    if sysname == "Windows":
        return (["powershell", "-NoProfile", "-NonInteractive",
                 "-Command", WIN_TOAST_PS],
                {"CC_TOAST_TITLE": title, "CC_TOAST_BODY": body})
    if sysname == "Darwin":
        script = (f'display notification "{body}" '
                  f'with title "{title}" sound name "Glass"')
        return (["osascript", "-e", script], {})
    if sysname == "Linux":
        return (["notify-send", title, body], {})
    return None


def fire_native_notify(title: str, body: str) -> None:
    if NOTIFY_FILE:  # テスト用: 実発火の代わりにファイルへ記録
        try:
            with open(NOTIFY_FILE, "a", encoding="utf-8") as f:
                f.write(f"{title}|{body}\n")
        except OSError:
            pass
        return
    cmd = native_notify_cmd(title, body)
    if not cmd:
        return
    argv, extra_env = cmd
    # interactive=True: トーストは対話セッションで起動しないと配信されない
    spawn_detached(argv, extra_env, interactive=True)  # 失敗してもフックを失敗させない
