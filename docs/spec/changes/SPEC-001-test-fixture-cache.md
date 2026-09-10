# SPEC-001: テスト用フィクスチャキャッシュ層の新設

<!-- 変更要求。承認後は凍結し、以後編集しない。現行仕様は docs/spec/features/ 側が持つ。 -->

| 項目 | 内容 |
|---|---|
| 対象FEAT | 新規(本SPECで FEAT-001 を新設) |
| 対象 | VBA |
| 起案日 | 2026-09-10 |

## 1. 目的・背景

`xlflow test` で外部Excelブックを開いて内部データ構造へ展開する入力処理が遅く、
テストの反復速度を律速している。外部ブックの内容はめったに変わらないため、
展開結果を再利用して2回目以降のロードを省く。

## 2. 要求事項

- REQ-1: 同一の外部ブックから同一のフィクスチャを2回以上ロードするとき、2回目以降は
  `Workbooks.Open` / `GetObject` を発行しない。
- REQ-2: キャッシュ対象は「外部ブックを読んだ生データ」ではなく、**突合せ・再設定まで
  済んだ最終状態**とする。加工処理のコストも省くのが目的であるため。
- REQ-3: キャッシュは2層とする。L1 = Excel プロセス常駐(メモリ)、L2 = ディスク永続。
  L2 は `xlflow push` によるモジュール再インポートと Excel プロセス再起動をまたいで有効。
- REQ-4: L1 はクラスオブジェクト(および `Collection` / `Dictionary`)をそのまま保持する。
  L2 は2次元配列(Variant)にフラット化して永続する。L2 ヒット時はクラス組み立てのみ実行する。
- REQ-5: 外部ブックが無変更でも**加工ロジックを変更した場合はキャッシュが無効化される**手段を
  提供する。フィクスチャローダーごとにスキーマ版定数を持ち、キャッシュキーに含める。
- REQ-6: 外部ブックの読み取り方式は `Workbooks.Open` を用い、読み取り時の Excel 状態
  (再計算・イベント・画面更新)を抑止する。ADO/OLEDB や xlsx 直パースは採らない。
- REQ-7: 本機構はテスト実行時のみ使用する。本番モジュールからは参照しない。

### スコープ外(この変更では触らない)

- 本番コード側の入力処理。本 SPEC は `src/modules/Tests/` 配下のみを追加する。
- L2 キャッシュファイルの世代管理・自動削除(溜まったら `ClearDisk` を手で叩く)。
- 具体的な業務フィクスチャのローダー実装。本 SPEC は**キャッシュ基盤とその自己テスト**まで。
- 読み取り方式の変更(ADO + ACE OLEDB / xlsx 直パース)。REQ-6 のとおり不採用。

## 3. 変更内容

現行に対応する機能は存在しない。以下を新規追加する。

### 3.1 参照順

```
テストコード
  └ TestFixture<名前>.Load…()                 フィクスチャ別ローダー(本SPECのスコープ外)
      ├ TestFixtureCache.TryGetObject(key)    L1: 加工済みオブジェクト → ヒットで終了
      ├ TestFixtureCache.TryGetArray(key)     L1配列 → L2ディスク の順に探索
      └ どちらもミス: 外部ブックを開いて加工 → PutArray → PutObject
```

### 3.2 追加モジュール

| モジュール | 配置 | 役割 | 概算行数 |
|---|---|---|---|
| `TestFixtureCache.bas` | `src/modules/Tests/` | L1/L2 の格納・取得、キー生成、無効化、破棄 | 約170 |
| `TestFixtureCodec.bas` | `src/modules/Tests/` | 2次元配列 ⇔ テキストの直列化・復元 | 約130 |
| `TestFixtureCacheTests.bas` | `src/modules/Tests/` | §6 の受け入れテスト | 約200 |

合計約500行。直列化をキャッシュ制御から分けるのは、直列化が独立して単体テスト可能な
関心事であり、混ぜると型の追加とキャッシュ方針の変更が同じファイルに集まるため。

### 3.3 公開 API

```vb
' キー生成。srcPath が存在しないときはエラーを発生させる
Public Function BuildKey(ByVal srcPath As String, ByVal loaderName As String, _
                         ByVal schemaVersion As Long) As String

' L1(メモリ)。クラスオブジェクトをそのまま保持する
Public Function TryGetObject(ByVal cacheKey As String, ByRef outObj As Object) As Boolean
Public Sub PutObject(ByVal cacheKey As String, ByVal obj As Object)

' L1配列 → L2ディスク の順に探索。2次元配列のみ
Public Function TryGetArray(ByVal cacheKey As String, ByRef outData As Variant) As Boolean
Public Sub PutArray(ByVal cacheKey As String, ByRef data As Variant)

' 破棄
Public Sub ClearMemory()   ' L1 のみ
Public Sub ClearDisk()     ' L2 を削除し、L1 も空にする(復旧手段)
Public Function CacheDir() As String
```

### 3.4 キャッシュキーと無効化

```
cacheKey = loaderName & "|" & LCase$(<絶対パス>) & "|" _
         & Format$(FileDateTime(path), "yyyymmddhhnnss") & "|" _
         & CStr(FileLen(path)) & "|v" & CStr(schemaVersion)
```

- `schemaVersion` は各ローダーが `Private Const SCHEMA_VERSION As Long = 1` として持ち、
  **加工ロジックを変更したら値を1つ上げる**。上げ忘れると古い加工結果でテストが通る
  (偽の緑)ため、ローダーの Doxygen ヘッダにこの規約を明記する。
- 上げ忘れの保険として `ClearDisk` を提供する。

### 3.5 L2 の配置と形式

- ディレクトリ: `tmp/test-fixture-cache/`。`tmp/` は `.gitignore` 済みのため、
  キャッシュが git に混入する経路が構造的に存在しない。
- ファイル名: `<loaderName>-<hash8>.txt`。`hash8` は `cacheKey` の FNV-1a 32bit を16進8桁。
  ハッシュ衝突は本文1行目のフルキー照合で検出し、不一致ならミス扱いとする。
- 文字コード: UTF-8(`CreateObject("ADODB.Stream")` による遅延バインド)。CP932 では
  外部マスタの機種依存文字が落ちるため。ADODB.Stream は Windows 標準構成要素で、
  ACE OLEDB のような別途インストールを要しない。
- 本文書式:

```
XLFCACHE<TAB>1<TAB><cacheKey>
<行数><TAB><列数><TAB><行下限><TAB><列下限>
<型タグ><TAB><値>          ← 行優先で 行数×列数 レコード
```

- 型タグ: `D`=Double(10進表記) / `H`=Double(16進16文字のビット表現) / `S`=String /
  `B`=Boolean / `E`=Empty / `N`=Null / `R`=エラー値(値は `CLng(CVErr)`)。
  `Range.Value2` は日付をシリアル値 Double で返すため `Date` 型は現れない。
  上記以外の `VarType` が来た場合は**エラーを発生させる**(無言で丸めない)。
  Double は10進で往復できる値のみ `D`、できない値は `H` とする(§4 No.11)。
- 文字列のエスケープ: `\`→`\\`、TAB→`\t`、CR→`\r`、LF→`\n`。

### 3.6 外部ブックを読むときの Excel 状態(REQ-6)

キャッシュミス時の実ロードでは以下を設定し、`Finally` 相当のブロックで必ず復元する。

| 設定 | 値 | 復元 |
|---|---|---|
| `Application.Calculation` | `xlCalculationManual` | 元の値へ戻す |
| `Application.EnableEvents` | `False` | 元の値へ戻す |
| `Application.ScreenUpdating` | `False` | 元の値へ戻す |
| `Workbooks.Open` の `ReadOnly` | `True` | — |
| `Workbooks.Open` の `UpdateLinks` | `0` | — |

セルの読み取りは `arr = ws.Range(...).Value2` の一括取得とし、逐次アクセスを禁じる。
COM 参照の解放は `.github/skills/xlflow/references/testing.md` の
「COM Object Cleanup in Tests」の手順に従う(`wb.Close False` の直後に `Set wb = Nothing`)。

## 4. 異常系・境界値

| No | 条件 | 期待動作 |
|---|---|---|
| 1 | `BuildKey` の `srcPath` が存在しない | `vbObjectError + 601` を `TestFixtureCache.BuildKey` を Source として発生させる |
| 2 | L2 ファイルの1行目が `XLFCACHE` 形式でない | `TryGetArray` は False を返す。当該ファイルを削除し、`XlflowDebug.Log` に破損理由を記録する。テストは失敗させない |
| 3 | L2 ファイルの1行目のキーが要求キーと不一致(ハッシュ衝突) | No.2 と同じ扱い |
| 4 | L2 のレコード数が宣言された 行数×列数 と一致しない | No.2 と同じ扱い |
| 5 | 型タグが `D/H/S/B/E/N/R` 以外、または `H` が16進16文字でない | No.2 と同じ扱い |
| 6 | `PutArray` に2次元でない値(スカラー・1次元配列・オブジェクト)を渡す | `vbObjectError + 602` を発生させる |
| 7 | 元ファイルが1秒以内に更新され、更新後もバイト数が同一 | `FileDateTime` は秒精度のため検出できない。**既知の制約**として受け入れる(外部マスタが1秒以内に2回変わる運用は無いという前提) |
| 8 | `tmp/test-fixture-cache/` が存在しない | 初回 `PutArray` 時に作成する |
| 9 | L2 への書き込みが失敗(権限・ディスク不足) | `XlflowDebug.Log` に記録し、L1 のみで処理を続行する。テストは失敗させない |
| 10 | `PutArray` に未初期化の動的配列(0行相当)を渡す | `vbObjectError + 602` を発生させる |
| 11 | 15桁の10進表記で往復しない Double(`0.1 + 0.2` / `1 / 3` / 表現範囲の境界値) | `H` のビット表現で保存し、復元時に元の値と1ビットも違わない |

<!-- No.11 追加(2026-09-10、実装着手後):
     VBA の Str$ / CStr / Format$ はいずれも有効15桁までで、10進テキストでは
     Double を無損失に往復できないことが実測で判明した(0.1+0.2 が 0.3 に化ける)。
     キャッシュヒット時だけ値が変わる=偽の緑になるため、代替案をユーザに提示し
     「型タグ H を追加するハイブリッド」を選択して承認を得たうえで追加した。
     凍結済み SPEC の改訂だが、AI の独断ではなく人間の裁定によるものである。 -->

<!-- No.10 改訂(2026-09-10、実装着手後):
     当初「行数0として保存し、復元時も0行の配列を返す」としていたが、VBA に
     要素数0の2次元配列は存在せず実現不可能であることが実測で判明した
     (`ReDim a(1 To 0, 1 To 3)` / `ReDim a(0 To -1, 0 To 2)` はいずれも実行時エラー9)。
     代替案をユーザに提示し、「空範囲はキャッシュせずエラーにする」を選択して
     承認を得たうえで本行へ差し替えた。凍結済み SPEC の改訂だが、AI の独断ではなく
     人間の裁定によるものであることを記録として残す。 -->


<!-- ClearDisk の役割改訂(2026-09-10、セルフレビュー後):
     当初は「L2 のみを削除する」と定めていたが、それでは同一 Excel プロセス内で
     L1 が古い加工結果を返し続け、SCHEMA_VERSION 上げ忘れの復旧手段として
     成立しないことがレビューで判明した。ユーザの裁定を得て L1 も空にする
     仕様へ改めた。凍結済み SPEC の改訂だが、AI の独断ではない。 -->

## 5. 影響範囲

- 追加: `src/modules/Tests/TestFixtureCache.bas`、`src/modules/Tests/TestFixtureCodec.bas`、
  `src/modules/Tests/TestFixtureCacheTests.bas`(§3.2、合計約500行)。
- 追加: `docs/spec/features/FEAT-001-test-fixture-cache.md`(§7 の内容で新規作成)。
- 変更なし: 既存の `XlflowAssert.bas` / `XlflowDebug.bas` / `XlflowRuntime.bas` /
  `XlflowUI.bas`、`.gitignore`(`tmp/` は既に無視対象)、`xlflow.toml`。
- 本番モジュールへの変更は無い(REQ-7 / スコープ外)。

**規模について**: 合計約500行で、AGENTS.md の「変更が200行を超える場合は着手前に規模と
方針を提示する」に該当する。本節をもって提示とし、承認をもって着手可とする。

## 6. 受け入れ基準

<!-- scripts/check_acceptance.py が検査する節。書式を崩さない。
     判定は「自動」「人手」のいずれか。空欄・その他は ERROR。
     自動: 検証欄に allowlist 内のコマンドを書く。実行はエージェントが
           Bash ツール経由で行う(検査スクリプトはコマンドを実行しない)。
           許可: xlflow test / xlflow lint / dotnet test / python scripts/ /
                 bash docs/spec/changes/ac/
     人手: 検証欄に bash docs/spec/changes/ac/SPEC-nnn-ACn.sh を書く。
           作り方は docs/spec/changes/ac/README.md。
     検証手段が思いつかない AC は、要求が曖昧である兆候。AC ではなく要求を直す。 -->

| AC | 内容 | 検証 | 判定 |
|---|---|---|---|
| AC-1 | 同一キーで2回ロードしたとき、実オープン回数カウンタが 1 のままである | `xlflow test --filter Test_Cache_SecondLoad_DoesNotOpenWorkbook --session --json` | 自動 |
| AC-2 | `ClearMemory` 実行後も `TryGetArray` が True を返し、実オープン回数カウンタが増えない | `xlflow test --filter Test_Cache_L2Hit_AfterClearMemory --session --json` | 自動 |
| AC-3 | 元ファイルの更新日時を変えると `TryGetArray` が False を返す | `xlflow test --filter Test_Cache_Invalidated_WhenSourceModified --session --json` | 自動 |
| AC-4 | `schemaVersion` を 1 から 2 に変えると `TryGetArray` が False を返し、旧データを返さない | `xlflow test --filter Test_Cache_Invalidated_WhenSchemaVersionChanged --session --json` | 自動 |
| AC-5 | Double / String / Boolean / Empty / エラー値 の5型が L2 往復後も `VarType` と値の両方で一致する | `xlflow test --filter Test_Codec_RoundTrip_PreservesTypes --session --json` | 自動 |
| AC-6 | TAB・CR・LF・`\` を含む文字列が L2 往復後に元の文字列と一致する | `xlflow test --filter Test_Codec_RoundTrip_EscapesControlChars --session --json` | 自動 |
| AC-7 | L2 ファイルの1行目を壊すと `TryGetArray` が False を返し、テストは失敗せず、`XlflowDebug` の出力に破損記録が残る | `xlflow test --filter Test_Cache_CorruptFile_FallsBackAndLogs --session --json` | 自動 |
| AC-8 | 存在しないパスで `BuildKey` を呼ぶと `vbObjectError + 601` が発生する | `xlflow test --filter Test_Cache_MissingSource_RaisesError --session --json` | 自動 |
| AC-9 | 2次元でない値を `PutArray` に渡すと `vbObjectError + 602` が発生する | `xlflow test --filter Test_Cache_PutArray_RejectsNon2DValue --session --json` | 自動 |
| AC-10 | L2 ヒットの所要時間が、同一データの実ロード所要時間の 50% 以下である | `xlflow test --filter Test_Cache_L2Hit_IsFasterThanRealLoad --session --json` | 自動 |
| AC-11 | `CacheDir` が返すパスがリポジトリ直下の `tmp\test-fixture-cache` を指す | `xlflow test --filter Test_Cache_CacheDir_IsUnderTmp --session --json` | 自動 |
| AC-12 | 追加した全モジュールの公開関数に Doxygen ヘッダがある | `python scripts/check_doxygen.py --scan src/modules/Tests` | 自動 |
| AC-13 | lint 指摘が 0 件である | `xlflow lint --json` | 自動 |

## 7. 正本への反映内容

<!-- check_spec_sync.py が機械転記・検査する節。書式を崩さない。
     見出しは「### FEAT-<番号> § <FEAT側の節見出し>」で、本文には
     **反映後の FEAT 該当節の完成形**をそのまま書く(差分や方針で書かない)。
     承認時にこの本文ごと承認し、実装後は FEAT の該当節をこの本文で置換する。
     節を複数変更する場合は ### を複数並べる。 -->

<!-- FEAT-001 は本SPECで新規作成する。メタ表は以下のとおり:
     | 対象 | VBA |
     | 対象モジュール | TestFixtureCache.bas, TestFixtureCodec.bas |
     | 関連SPEC | SPEC-001 |
     | 出典 | SPEC-001 | -->

### FEAT-001 § 1. 入力

テストコードから `TestFixtureCache` の公開関数を呼び出して使う。起動用のマクロは持たない。

| 引数 | 型 | 意味 |
|---|---|---|
| `srcPath` | String | 展開元となる外部ブックのパス。相対パスはブックの場所を基準に絶対化する |
| `loaderName` | String | フィクスチャローダーの識別名。キャッシュファイル名の接頭辞になる |
| `schemaVersion` | Long | ローダーの加工ロジックの版。ローダーが `Private Const SCHEMA_VERSION` として持つ |
| `cacheKey` | String | `BuildKey` の戻り値 |
| `data` | Variant | 2次元配列。`PutArray` の入力 |
| `obj` | Object | クラス / `Collection` / `Dictionary`。`PutObject` の入力 |

参照するファイルは `srcPath` と、`CacheDir` 配下の `<loaderName>-<hash8>.txt`。
参照するシート・設定値は無い。

### FEAT-001 § 2. 処理

2層のキャッシュを持つ。L1 は Excel プロセス常駐のメモリ、L2 はディスク上のテキストファイル。
L1 は `xlflow push` によるモジュール再インポートで消えるが、L2 は残る。

**`BuildKey(srcPath, loaderName, schemaVersion)`**

1. `srcPath` を絶対パス化する。ファイルが存在しなければ `vbObjectError + 601` を発生させる。
2. 次の文字列を組み立てて返す。

```
loaderName & "|" & LCase$(絶対パス) & "|" & Format$(FileDateTime(絶対パス), "yyyymmddhhnnss") _
  & "|" & CStr(FileLen(絶対パス)) & "|v" & CStr(schemaVersion)
```

更新日時とバイト数を含めるため元ファイルの変更でキーが変わる。`schemaVersion` を含めるため、
元ファイルが無変更でも加工ロジックの版を上げればキーが変わる。

**`TryGetObject(cacheKey, outObj)`**

L1 のオブジェクト辞書のみを探索する。存在すれば `outObj` に設定して True、無ければ False。

**`PutObject(cacheKey, obj)`**

L1 のオブジェクト辞書に格納する。ディスクへは書かない。

**`TryGetArray(cacheKey, outData)`**

1. L1 の配列辞書を探索する。ヒットすれば `outData` に代入して True を返す。
2. ミスなら `CacheDir` 配下の `<loaderName>-<hash8>.txt` を読む(`hash8` は `cacheKey` の
   FNV-1a 32bit を16進8桁)。ファイルが無ければ False。
3. 1行目が `XLFCACHE<TAB>1<TAB><cacheKey>` と完全一致しなければ、当該ファイルを削除し、
   `XlflowDebug.Log` に破損理由を記録して False を返す。
4. 2行目の 行数・列数・行下限・列下限 に従って本文レコードを復元する。レコード数が
   行数×列数 と一致しない場合、または型タグが `D/H/S/B/E/N/R` 以外の場合は手順3と同じ扱いとする。
5. 復元した配列を L1 の配列辞書へも格納したうえで `outData` に代入し、True を返す。

**`PutArray(cacheKey, data)`**

1. `data` が2次元配列でなければ `vbObjectError + 602` を発生させる。
2. L1 の配列辞書に格納する。
3. `CacheDir` が存在しなければ作成し、§3(出力)の書式で UTF-8 のテキストファイルへ書き出す。
   書き込みに失敗した場合は `XlflowDebug.Log` に記録し、L1 のみ有効な状態で正常終了する。

**`ClearMemory` / `ClearDisk`**

`ClearMemory` は L1 の両辞書を空にする。`ClearDisk` は `CacheDir` 配下のキャッシュファイルを
削除したうえで `ClearMemory` を呼び、L1 も空にする。ローダーの `SCHEMA_VERSION` を上げ忘れた
際の復旧手段であり、L1 を残すと同一 Excel プロセス内で古い加工結果を返し続けて復旧が
成立しないため、両層をまとめて破棄する。

**外部ブックを読む側の規約**

キャッシュミス時に外部ブックを読むローダーは、以下を設定し、エラー時も含めて必ず元の値へ復元する。

| 設定 | 値 |
|---|---|
| `Application.Calculation` | `xlCalculationManual` |
| `Application.EnableEvents` | `False` |
| `Application.ScreenUpdating` | `False` |
| `Workbooks.Open` の `ReadOnly` | `True` |
| `Workbooks.Open` の `UpdateLinks` | `0` |

セルの読み取りは `Range.Value2` による一括取得とし、セル単位の逐次アクセスは行わない。
`Workbooks.Open` / `GetObject` で得た参照は `wb.Close False` の直後に `Set wb = Nothing` で解放する。

### FEAT-001 § 3. 出力

| 関数 | 戻り値 |
|---|---|
| `BuildKey` | String。キャッシュキー |
| `TryGetObject` / `TryGetArray` | Boolean。ヒット時 True。ヒット時のみ `ByRef` 引数に値が入る |
| `PutObject` / `PutArray` / `ClearMemory` / `ClearDisk` | 無し(Sub) |
| `CacheDir` | String。`tmp\test-fixture-cache` の絶対パス |

副作用として `CacheDir` 配下に UTF-8 のテキストファイルを作成・削除する。書式は次のとおり。

```
XLFCACHE<TAB>1<TAB><cacheKey>
<行数><TAB><列数><TAB><行下限><TAB><列下限>
<型タグ><TAB><値>          ← 行優先で 行数×列数 レコード
```

型タグは `D`=Double(10進表記)、`H`=Double(16進16文字のビット表現)、`S`=String、
`B`=Boolean、`E`=Empty、`N`=Null、`R`=エラー値(値は `CLng(CVErr)`)。
`Range.Value2` は日付をシリアル値 Double で返すため `Date` 型は現れない。
文字列は `\`→`\\`、TAB→`\t`、CR→`\r`、LF→`\n` にエスケープする。

Double は原則 `D` の10進表記で書く。ただし VBA の `Str$` / `CStr` / `Format$` は
いずれも有効15桁までしか出せず、`0.1 + 0.2` や `1 / 3` は10進の往復で別の値になる。
`Trim$(Str$(value))` を `Val` で戻して元の値と一致しない場合(および丸めが Double の
表現範囲を超えて `Val` がオーバーフローする場合)は `H` のビット表現へ退避する。
金額・数量・日付シリアルなど実データの大半は `D` を通るため、キャッシュファイルは
目視で読める状態が保たれる。

`tmp/` は `.gitignore` の対象であり、キャッシュファイルが git に追跡される経路は無い。

### FEAT-001 § 4. 異常系・境界値

| No | 条件 | 期待動作 |
|---|---|---|
| 1 | `BuildKey` の `srcPath` が存在しない | `vbObjectError + 601` を Source `TestFixtureCache.BuildKey` で発生させる |
| 2 | L2 ファイルの1行目が `XLFCACHE` 形式でない | `TryGetArray` は False。当該ファイルを削除し `XlflowDebug.Log` に破損理由を記録する |
| 3 | L2 ファイルの1行目のキーが要求キーと不一致(ハッシュ衝突) | No.2 と同じ |
| 4 | L2 のレコード数が 行数×列数 と一致しない | No.2 と同じ |
| 5 | 型タグが `D/H/S/B/E/N/R` 以外、または `H` が16進16文字でない | No.2 と同じ |
| 6 | `PutArray` に2次元でない値を渡す | `vbObjectError + 602` を発生させる |
| 7 | 元ファイルが1秒以内に更新され、更新後もバイト数が同一 | `FileDateTime` が秒精度のため検出できない。既知の制約 |
| 8 | `CacheDir` が存在しない | 初回 `PutArray` 時に作成する |
| 9 | L2 への書き込みが失敗した | `XlflowDebug.Log` に記録し、L1 のみで続行する |
| 10 | `PutArray` に未初期化の動的配列(0行相当)を渡す | `vbObjectError + 602` を発生させる |
| 11 | 15桁の10進表記で往復しない Double(`0.1 + 0.2` / `1 / 3` / 表現範囲の境界値) | `H` のビット表現で保存し、復元時に元の値と1ビットも違わない |

### FEAT-001 § 5. この機能固有のルール

- 本機構はテスト専用である。本番モジュールから `TestFixtureCache` / `TestFixtureCodec` を
  参照しない。テストが本番と別経路を通るため、入力処理そのものの正しさは本機構を経由しない
  テストで別途担保する。
- キャッシュするのは「外部ブックを読んだ生データ」ではなく、突合せ・再設定まで済ませた
  最終状態である。
- L1 はクラスオブジェクトを保持し、L2 は2次元配列のみを永続する。L2 ヒット時は
  クラス組み立てのみを実行する。クラスのシリアライザは書かない。
- フィクスチャローダーは `Private Const SCHEMA_VERSION As Long` を持ち、**加工ロジックを
  変更したら値を1つ上げる**。上げ忘れると元ファイル無変更のままキャッシュがヒットし、
  古い加工結果でテストが通る(偽の緑)。この規約はローダーの Doxygen ヘッダに明記する。
- キャッシュの破損は落とさずミス扱いにするが、`XlflowDebug.Log` への記録を省略しない
  (記録しない場合は「エラーの握りつぶし」に当たる)。

## 8. 未決事項

- [ ] L2 キャッシュファイルの世代管理。当面は溜まり続け、`ClearDisk` を手動で叩く運用とする。
      件数上限や TTL が必要になったら別 SPEC で扱う。
- [ ] `SCHEMA_VERSION` の上げ忘れを機械検出する手段。現状は規約と code-review 頼み。
      ローダーのソースハッシュをキーに含める案があるが、VBA からソースを読む手段が
      素直でないため本 SPEC では採らない。
- [ ] AC-10 の 50% という閾値。実測を省略して起案しているため、根拠は「キャッシュが
      効いていることを検出できる最小限の差」でしかない。初回実装時の実測値を見て、
      閾値が緩すぎる/厳しすぎる場合は別 SPEC で調整する。
