# AI 教訓ログ

修復ループ・仕様矛盾・reviewer の Critical 指摘が発生したタスクでのみ、
解決後に1〜3行で追記する(何が起きたか / 根本原因 / 次回の回避策)。
問題が発生しなかったタスクでは記録しない。

## 2026-07-11 サブエージェント完了トーストが一切表示されなかった

- 何が起きたか: ハーネスv2.3導入後、A4/A5(Windowsトースト実表示)検証で、
  完了・途中経過トーストが一度も出ていなかった。detached起動でエラーが
  隠れ、切り分けに複数回の実測を要した。原因は3つ複合していた。
- 根本原因(3件、`scripts/subagent_common.py`):
  1. WinRT 型リテラル `[Type, Assembly, ContentType=WindowsRuntime]` を
     カンマ後に改行して2行に分割。PowerShellが「型名にアセンブリ名が
     指定されていません」で構文エラー → スクリプト全体が起動せず無表示。
  2. AUMID フォールバックの試行順が「未登録の 'Claude Code' が先頭」。
     未登録AUMIDは CreateToastNotifier/Show が**例外を出さずに表示だけ
     されない**(サイレント失敗)ため、例外ベースの foreach では
     `break` して登録済みAUMID(Windows PowerShell)に到達しなかった。
  3. Show() は非同期キューのみ。detached プロセスが直後に終了すると
     配信前に消える。加えて `DETACHED_PROCESS` 起動だと非対話セッション
     扱いになり、対話デスクトップにトーストが届かない。
- 修正: (1)型リテラルを1行化 (2)登録済みAUMIDを先頭に (3)末尾に
  `Start-Sleep 5` + 通知spawnを `DETACHED_PROCESS` → `CREATE_NO_WINDOW`
  (対話セッション維持)。watchdog本体のspawnは120秒生存が要るので
  `DETACHED_PROCESS` のまま分離。
- 次回の回避策: PowerShellの型リテラルは1行で書く。外部トースト等の
  best-effort/detachedコマンドは、導入時に必ず一度は同期実行
  (capture_output)で returncode/stderr を実測してサイレント失敗を潰す。
  「例外が出ない=成功」ではない(未登録AUMIDは無例外で無表示)。
  トーストは `CREATE_NO_WINDOW`(対話セッション)で起動する。
  `DETACHED_PROCESS` は生存が必要な監視プロセス専用。
