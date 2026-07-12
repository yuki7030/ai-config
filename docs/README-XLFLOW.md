# xlflow 導入ガイド

## 検討結果サマリ

xlflow (harumiWeb, MIT, v0.20.0 / 2026-07-08 時点で活発に開発中) は以下を解決する:

- **エラーダイアログの吸収** … 実行時/コンパイルエラーのGUIを自動で閉じ、
  エラー内容+発生行をターミナルにJSONで返す(自作ps1では実現不可だった)
- **MsgBox/InputBoxのヘッドレス化** … XlflowUIラッパー + `--msgbox id=yes`
- **セッションモード** … ブックを開いたままpush/run/testを高速に回せる
  (自作ps1は毎回Excelを開閉していた)
- **VBA LSP** … 補完・診断・定義ジャンプ・Rename・Test Explorer
- **UserFormのYAML定義(UaC)** … フォームもGit管理・AI生成可能
- **lint / fmt / inspect / formulas pull** … 静的解析からシート状態検証まで

PowerShell非依存(Go/C#製)、pull=UTF-8 / push=CP932自動変換も内蔵。

### リスク(許容範囲と判断)

- v0.x のため破壊的変更の可能性あり(CHANGELOGはよく整備されている)
- 個人開発(週次ペースでリリース中、直近コミットは前日)
- Windows + Excel + トラストセンター設定必須(現環境では問題なし)

## 構築手順

以下の手順1〜5(トラストセンター設定を除く)は [scripts/xlflow-setup.ps1](../scripts/xlflow-setup.ps1) で
完全自動実行できる。新規/既存の別はスクリプト実行時にターミナルで確認される。

```powershell
powershell -ExecutionPolicy Bypass -File scripts\xlflow-setup.ps1
```

### 1. xlflow本体のインストール(いずれか)

```powershell
winget install HarumiWeb.Xlflow
# または
irm https://harumiweb.github.io/xlflow/install.ps1 | iex
```

確認: `xlflow --version` / 環境診断: `xlflow doctor --json`

### 2. Excelの設定(未設定の場合)

トラストセンター → マクロの設定 →
「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」を ON

### 3. VSCode拡張のインストール

```powershell
code --install-extension harumiWeb.xlflow-vscode
```

### 4. プロジェクト作成

新規:
```powershell
xlflow new --with-skill
```

既存ブック(例: vba-ai-test-kit で使っていた xlsm):
```powershell
xlflow init excel_file\Book.xlsm --agent agents --with-skill
```

`--agent agents` を指定するとエージェント選択画面を出さず `.agents/skills` へ導入される
(scripts/xlflow-setup.ps1 はこれを自動実行する)。

### 5. Agent Skill を Claude Code / Copilot 両対応にする

xlflow のスキルインストール先は以下から選択できる:
`.agents/skills`(共有) / `.codex/skills` / `.claude/skills` / `.cursor/skills` / `.gemini/skills`

本リポジトリは他スキル(vba-coding 等)と同様、**実体を `.github/skills/xlflow/` に集約**し、
`.claude/skills/xlflow` と `.agents/skills/xlflow` はそこへの**シンボリックリンク**にしている
(重複を持たず、Claude Code / agents 双方から live スキルとして参照可能)。
scripts/xlflow-setup.ps1 は `.agents/skills/xlflow` を `.github/skills/xlflow` へ移動した上で
両方向のシンボリックリンクを自動で張る。

xlflow 開発ループと規則は AGENTS.md 内の「Excel VBA 開発 (xlflow)」節に記載済み。
Claude Code / Copilot の双方から `.github/skills/xlflow/SKILL.md` を参照させる。

### 6. 動作確認

```powershell
xlflow session start
xlflow run --json          # project.entry (Main.Run) が実行される
xlflow test
xlflow session stop
```

### 7. AIエージェントへの指示例

```
/xlflow 指定フォルダ内の請求書Excelを集計するマクロをTDDで実装して。
- テストを先に書き、xlflow test が通るまで修正
- 実装後に xlflow lint と xlflow fmt を実行
- 最後に xlflow push でブックへ反映
```

## xlflow.toml 設定リファレンス

`xlflow.toml` はプロジェクトルート直下にあり、xlflow各コマンドの挙動を決める唯一の設定ファイル。
値自体はファイル内コメントで定義済みなので、ここでは「何を」「いつ」直すかの判断観点をまとめる。

### [project] — プロジェクト識別・エントリポイント
| キー | 既定値 | 内容 | 修正するタイミング |
|---|---|---|---|
| name | "vba-tool" | 出力メッセージ用のプロジェクト名 | リポジトリ名/ブック名を変更した時のみ |
| entry | "Main.Run" | `xlflow run`(マクロ省略時)で実行される既定マクロ | エントリーポイントのモジュール/Sub名を変更した時 |

### [excel] — ワークブック接続・自動化設定
| キー | 既定値 | 内容 | 修正するタイミング |
|---|---|---|---|
| path | "build/vba-tool.xlsm" | ビルド対象ブックのパス | ビルド出力先を変えたい時(通常不要) |
| visible | false | Excelウィンドウ表示 | 手元でUI挙動を目視デバッグしたい時だけ一時的にtrue |
| display_alerts | false | 上書き確認等のExcelアラート抑制 | 基本falseのまま維持。自動化が想定外のダイアログで止まる原因調査時のみ一時的にtrue |
| bridge | "auto" | Excel COMブリッジの接続方式 | auto判定で接続が不安定/失敗する環境でdotnet固定に切り替え |

### [src] — ソースディレクトリ構成
`modules` / `classes` / `forms` / `workbook` の4キーで `src/` 配下のマッピングを定義。
ディレクトリ構成そのものを変更しない限り基本触らない。

### [vba] — VBEコンポーネントのフォルダ注釈(Rubberduckスタイル)
| キー | 既定値 | 内容 | 修正するタイミング |
|---|---|---|---|
| folders | true | `@Folder("A.B")` 注釈と機能別フォルダ配置を有効化 | フラットな一覧管理に戻したい時のみfalse |
| folder_annotation | "update" | pushのたびにソースディレクトリ構成から注釈を書き換えるか | VBE側で手動整理した注釈を保持したいならpreserve。注釈自体を使わないならignore |
| default_component_folders | true | 新規モジュール追加時に配置パスから自動でフォルダ注釈を付与 | 手動でフォルダ注釈を管理したい場合false |

### [userform] — UserFormコードの置き場所
`code_source`: `"frm"`(既定・コードは.frm内)/ `"sidecar"`(`src/forms/code/<FormName>.bas` に分離)。
フォーム差分のレビューしやすさを優先する場合はsidecarへの切り替えを検討する。

### [fmt] — `xlflow fmt` の整形挙動
`operator_spacing` / `declaration_spacing` はいずれも既定true。
チームの整形方針が既定と異なる場合のみfalseに変更する。

### [lint] — 静的解析(pushのpreflightにも使用)
`disabled_rules` に診断ID(例: `"VB020"`)を追加するとルール単位で無効化できる
(v0.18以降 `disabled_rules = ["VB020"]` 形式に対応)。

**修正するタイミング**: 日本語プロジェクト特有の誤検知が続く時、または意図的な設計判断
(例: レガシーコードでPublicフィールドを許容)の時。
むやみな無効化はAGENTS.mdの「エラーの握りつぶし禁止」に抵触しうるため、
無効化する場合は理由をコメントで残してから追加すること。

VB020(未使用ローカル変数)は既定有効。VB018(スコープシャドウイング)/
VB021(未使用Privateプロシージャ)/VB027(With曖昧参照)は既定コメントアウト(オプトイン)。

### [analyze] — 実行時リスク解析(`xlflow analyze`)
`disabled_rules` で実行時リスク診断(例: VBA205 ActiveSheet依存)を個別無効化できる。

**修正するタイミング**: レガシー資産の移行中など、意図的に許容しているリスクパターンに
誤検知が続く時のみ。新規実装では極力ルールに従い、解析結果自体をコードで解消する方を優先する。

### 全体の判断基準
- `xlflow.toml` の緩和設定は最終手段。まず `src/` 側を規約に合わせることを優先する
- CI/自動化が想定外のダイアログ・上書き確認で止まる → `[excel] display_alerts` / `visible`
- push時のlintで日本語や意図的な設計が誤検知される → `[lint] disabled_rules`
- 実行時リスク解析でレガシー依存が大量検出される → `[analyze] disabled_rules`
- フォーム差分が大きくレビューしづらい → `[userform] code_source = "sidecar"`

## 運用メモ

- 反復中は `xlflow push --fast --session --no-save --json` が高速
- push時にlintがpreflight実行される(VB001〜: Option Explicit必須、
  .Select/.Activate禁止、暗黙Variant検出など)。誤検出時の対処は上記
  [lint] 節を参照
- デバッグ出力は Debug.Print ではなく XlflowDebug.Log を使う
  (run/test 実行時にターミナルへストリームされる)
- ダイアログは XlflowUI.MsgBox("id", ...) 形式で書く。通常実行時は
  普通のMsgBoxとして動くため、利用者向けの挙動は変わらない
