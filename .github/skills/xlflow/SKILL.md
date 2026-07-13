---
name: xlflow
description: このリポジトリの Excel VBA 開発(xlflow)の手順・安全規則。VBAの実装・テスト・push・シート検証を行う際に必ず使用。VBEは使わず xlflow CLI で回す。
---

# xlflow 開発ガイド

このリポジトリの VBA 開発は xlflow で行う。VBE は使用しない。
`src/` 配下は UTF-8 のまま編集し、CP932 変換は push 時に自動で行われる。

## 基本ループ

1. `xlflow session start` でセッションを開始する
2. `src/` 配下の .bas / .cls を編集する(UTF-8のまま。CP932変換は
   push時に自動で行われる)
3. `xlflow push --fast --session --no-save --json` で反映する
   (preflightのlintで失敗したら指摘箇所を修正して再push)
4. `xlflow run --session --json` または `xlflow test` で実行・検証する
5. エラーはJSONの Diagnostic (種別・モジュール・行番号・コード) を読んで
   自己修正する。人間への確認は不要
6. 完了したら `xlflow save --json` → `xlflow session stop`

## 規則

- `Debug.Print` 禁止 → `XlflowDebug.Log` を使う
- `MsgBox` / `InputBox` 直接使用禁止 → `XlflowUI.MsgBox("dialog-id", ...)`
  形式を使い、実行時は `--msgbox dialog-id=yes` 等で応答を与える
- テストは `src/modules/Tests/` に `Test` 接頭辞のSubで書き、
  `XlflowAssert` (AssertEquals / AssertNotEqual / AssertTrue / AssertFalse /
  AssertIsNothing / AssertIsNotNothing / AssertFail / AssertInconclusive)
  でアサーションする。`BeforeAll` / `BeforeEach` / `AfterEach` / `AfterAll`
  フックと `'@Tag("...")` が使える
- 新機能はTDDで実装する: テスト作成 → `xlflow test` 失敗確認 → 実装 →
  PASS → `xlflow lint` → `xlflow fmt . --write`
- UserFormは .frm を直接編集せず `src/forms/specs/*.yaml` を編集して
  `xlflow form build <spec> --session --json` で生成する
- シート状態の検証は `xlflow inspect range --sheet <名前> --address <範囲> --json`
- ワークシート数式に依存する変更の前には `xlflow formulas pull --json` で
  数式スナップショットを確認する

## 関連

- VBA コーディング規約(命名・Doxygen・エラー処理) →
  [.github/skills/vba-coding/SKILL.md](../vba-coding/SKILL.md)
- 導入・セットアップ手順(setup.ps1・リンク構造・トラストセンター) →
  [docs/README-XLFLOW.md](../../../docs/README-XLFLOW.md)
