# =====================================================================
# xlflow-setup.ps1 — xlflow 環境セットアップ(完全自動)
# 実行: powershell -ExecutionPolicy Bypass -File scripts\xlflow-setup.ps1
#       [-Workbook path\to\Book.xlsm]   # 既存ブックのパスを事前指定する場合(省略時は対話で確認)
# =====================================================================
param(
    [string]$Workbook = ""
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

function Step($msg) { Write-Host "`n== $msg ==" -ForegroundColor Cyan }
function Info($msg) { Write-Host $msg -ForegroundColor Gray }
function Warn($msg) { Write-Host $msg -ForegroundColor Yellow }

function Update-SessionPath {
    # winget install直後でもターミナルを開き直さずxlflowコマンドを使えるようにする
    $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $user = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machine;$user"
}

function Set-SkillLink($LinkPath, $RelativeTarget) {
    if (Test-Path $LinkPath) {
        $item = Get-Item $LinkPath -Force
        if ($item.LinkType) {
            Info "$LinkPath は既にリンク済みのためスキップします"
            return
        }
        Warn "$LinkPath は実体ディレクトリのため削除してリンクに置き換えます"
        Remove-Item $LinkPath -Recurse -Force
    } else {
        New-Item -ItemType Directory -Force -Path (Split-Path $LinkPath -Parent) | Out-Null
    }
    try {
        New-Item -ItemType SymbolicLink -Path $LinkPath -Value $RelativeTarget -ErrorAction Stop | Out-Null
        Info "$LinkPath → $RelativeTarget のリンクを作成しました"
    } catch {
        Warn "シンボリックリンク作成に失敗($LinkPath)。開発者モードを有効化するか管理者権限で再実行してください: $_"
    }
}

# --- 0. プロジェクト種別の確認(対話) ---
Step "プロジェクト種別の確認"
if (-not $Workbook) {
    Write-Host "  [1] 新規プロジェクト(新しいマクロブックを作成)"
    Write-Host "  [2] 既存ブックから作成(既存の .xlsm を xlflow プロジェクト化)"
    do {
        $choice = Read-Host "選択してください (1 または 2)"
    } while ($choice -ne "1" -and $choice -ne "2")

    if ($choice -eq "2") {
        do {
            $Workbook = (Read-Host "既存ブックのパスを入力してください").Trim('"')
            if (-not (Test-Path $Workbook)) {
                Warn "ファイルが見つかりません: $Workbook"
            }
        } while (-not (Test-Path $Workbook))
    }
}

# --- 1. xlflow本体 ---
Step "xlflow のインストール確認"
if (-not (Get-Command xlflow -ErrorAction SilentlyContinue)) {
    Write-Host "xlflow が見つかりません。winget でインストールします..."
    winget install --id HarumiWeb.Xlflow -e --accept-source-agreements --accept-package-agreements --disable-interactivity
    Update-SessionPath
    if (-not (Get-Command xlflow -ErrorAction SilentlyContinue)) {
        Warn "インストール直後のためPATHが反映されていません。ターミナルを開き直して再実行してください"
        exit 1
    }
}
xlflow --version

# --- 2. 環境診断 (トラストセンター設定の検出を含む) ---
Step "xlflow doctor による環境診断"
xlflow doctor
Info "NG項目があれば docs/README-XLFLOW.md 手順2 (トラストセンター設定) を確認してください"

# --- 3. VSCode拡張 ---
Step "VSCode拡張 (harumiWeb.xlflow-vscode) のインストール"
if (Get-Command code -ErrorAction SilentlyContinue) {
    code --install-extension harumiWeb.xlflow-vscode --force
    # Barretta と併存させない (言語関連付け競合の回避)
    $barretta = code --list-extensions | Select-String -Pattern "vscode-barretta"
    if ($barretta) {
        Warn "Barretta 拡張が検出されました。無効化を推奨します (MIGRATION.md 参照)"
    }
} else {
    Warn "code コマンドが見つかりません。VSCodeから手動でインストールしてください"
}

# --- 4. プロジェクト初期化 (エージェント選択画面・更新確認プロンプトを --agent / --no-update-check で抑止) ---
Step "プロジェクト初期化"
if ($Workbook) {
    xlflow init $Workbook --agent agents --with-skill --with-module --no-update-check --json
} else {
    xlflow new --agent agents --with-skill --no-update-check --json
}

# --- 5. Claude Code / GitHub Copilot 共存設定 ---
# 現行運用: スキル実体は .github/skills/xlflow に一元化し、
# .claude/skills/xlflow と .agents/skills/xlflow はそこへのシンボリックリンクにする
# (重複を持たずClaude Code / Copilot 双方からliveスキルとして参照できるようにするため)
Step "Claude Code / Copilot スキル共有設定"
$githubSkillDir = ".github/skills/xlflow"
$agentsSkillDir = ".agents/skills/xlflow"
$claudeSkillDir = ".claude/skills/xlflow"

if (-not (Test-Path $githubSkillDir)) {
    if (Test-Path $agentsSkillDir) {
        New-Item -ItemType Directory -Force -Path ".github/skills" | Out-Null
        Move-Item $agentsSkillDir $githubSkillDir
        Info "スキル実体を $agentsSkillDir → $githubSkillDir へ移動しました"
    } else {
        Warn "$agentsSkillDir が見つからないため共有設定をスキップしました"
    }
}
if (Test-Path $githubSkillDir) {
    Set-SkillLink $agentsSkillDir "../../.github/skills/xlflow"
    Set-SkillLink $claudeSkillDir "../../.github/skills/xlflow"
}

# --- 6. 次の手順 ---
Step "セットアップ完了。動作確認コマンド"
Write-Host @"
  xlflow session start
  xlflow run --json
  xlflow test
  xlflow session stop
"@
