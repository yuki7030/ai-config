# ログ追尾タスクの実体。
#
# VS Code のタスクは folderOpen でしか自動起動できず、統合ターミナルは
# 外部プロセス(フック等)からは開けない。そこでタブ自体は folderOpen で
# 開いておき、実際の追尾開始は -Trigger の出現まで待つ。
# これで「起動時から古いログが流れる」ことを避けつつ統合ターミナルを保つ。
#
# トリガ:
#   xlflow -> .xlflow/session.json (session start で作られ stop で消える)
#   Claude -> .claude/logs/.session-active (SessionStart/SessionEnd フック)
#
# 【重要】このファイルは UTF-8 BOM付き で保存すること。
# Windows PowerShell 5.1 は BOM が無いスクリプトを CP932 として読むため、
# BOM を落とすと日本語コメントでパースエラーになる。
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Trigger,

    [string]$Label = 'session',

    # 指定すると XlflowDebug.Log 行の接頭辞を削らず生のまま出す
    [switch]$Raw,

    # xlflow --json の出力ブロックを畳む。成功時は1行に要約し、
    # 失敗時(status=failed / error 非 null)は全文をそのまま出す
    [switch]$HideJson
)

# ログ本文に日本語(VBA のエラーメッセージ等)が混ざるため、
# コンソール側の出力エンコーディングを UTF-8 に合わせる
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$pollMs = 250

# "xlflow: debug source=XlflowDebug.Log mode=interactive message=<本文>" の
# 接頭辞は毎行同じで読みづらいため、表示時だけ削る(ログファイルは生のまま)。
# 最初の message= までを非貪欲に食うので、本文に message= があっても壊れない。
$debugPrefix = '^xlflow: debug .*?message='

# JSON ブロックを畳んでいる間の蓄積。$null なら非蓄積中
$script:jsonBuffer = $null
$script:jsonDepth = 0

function Get-JsonField([string]$text, [string]$name) {
    $m = [regex]::Match($text, '"' + [regex]::Escape($name) + '"\s*:\s*"([^"]*)"')
    if ($m.Success) { return $m.Groups[1].Value }
    $m = [regex]::Match($text, '"' + [regex]::Escape($name) + '"\s*:\s*([0-9.]+)')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function Flush-JsonBuffer {
    if ($null -eq $script:jsonBuffer) { return }
    $text = $script:jsonBuffer -join "`n"
    $script:jsonBuffer = $null
    $script:jsonDepth = 0

    $status = Get-JsonField $text 'status'
    $failed = ($status -and $status -ne 'ok') -or ($text -match '"error"\s*:\s*\{')

    if ($failed) {
        # 失敗の詳細は畳まない。ここを削ると原因が追えなくなる
        foreach ($l in ($text -split "`n")) { Write-Host $l -ForegroundColor Red }
        return
    }

    $parts = @()
    foreach ($name in @('command', 'status')) {
        $v = Get-JsonField $text $name
        if ($v) { $parts += "$name=$v" }
    }
    $macro = Get-JsonField $text 'name'
    if ($macro) { $parts += "macro=$macro" }
    $ms = Get-JsonField $text 'duration_ms'
    if ($ms) { $parts += "${ms}ms" }

    Write-Host ("[json] " + ($parts -join ' ')) -ForegroundColor DarkGray
}

function Write-LogLine([string]$line) {
    if ($HideJson) {
        # 文字列リテラルを除いてから波括弧を数える(パス等の { で誤検知しないため)
        $stripped = $line -replace '"(\\.|[^"\\])*"', ''
        $open = ([regex]::Matches($stripped, '\{')).Count
        $close = ([regex]::Matches($stripped, '\}')).Count

        if ($null -eq $script:jsonBuffer) {
            if ($line.TrimStart().StartsWith('{')) {
                $script:jsonBuffer = @($line)
                $script:jsonDepth = $open - $close
                if ($script:jsonDepth -le 0) { Flush-JsonBuffer }
                return
            }
        }
        else {
            $script:jsonBuffer += $line
            $script:jsonDepth += ($open - $close)
            if ($script:jsonDepth -le 0) { Flush-JsonBuffer }
            return
        }
    }

    if (-not $Raw -and $line -match $debugPrefix) {
        Write-Host ($line -replace $debugPrefix, '')
    }
    else {
        Write-Host $line
    }
}

while ($true) {
    Write-Host "waiting for $Label session..." -ForegroundColor DarkGray
    while (-not (Test-Path -LiteralPath $Trigger)) {
        Start-Sleep -Milliseconds $pollMs
    }

    # 追尾対象がまだ無い状態でも待機できるよう、空ファイルだけ用意する。
    # 既存ファイルは -Force で切り詰められるため触らない。
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType File -Force -Path $Path | Out-Null
    }

    Write-Host ("=== {0} session start {1} ===" -f $Label, (Get-Date -Format 'HH:mm:ss')) -ForegroundColor Cyan

    # Get-Content -Wait は制御を返さずトリガの消滅を検知できないため、
    # ストリームを自前で持ち、セッション開始時点の末尾から読む
    $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    [void]$reader.BaseStream.Seek(0, 'End')

    while (Test-Path -LiteralPath $Trigger) {
        # ログが作り直された(サイズが読み取り位置より小さい)場合は先頭へ戻す
        if ($reader.BaseStream.Length -lt $reader.BaseStream.Position) {
            [void]$reader.BaseStream.Seek(0, 'Begin')
            $reader.DiscardBufferedData()
        }

        $line = $reader.ReadLine()
        if ($null -ne $line) {
            Write-LogLine $line
        }
        else {
            Start-Sleep -Milliseconds $pollMs
        }
    }

    # セッション終了直前に書かれた分を取りこぼさない
    while ($null -ne ($line = $reader.ReadLine())) {
        Write-LogLine $line
    }
    # 途中で切れた JSON ブロックを抱えたまま終わらない
    Flush-JsonBuffer

    $reader.Dispose()
    $stream.Dispose()
    Write-Host ("=== {0} session end {1} ===" -f $Label, (Get-Date -Format 'HH:mm:ss')) -ForegroundColor DarkGray
}
