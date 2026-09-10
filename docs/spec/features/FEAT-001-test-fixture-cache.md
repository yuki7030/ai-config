# FEAT-001: テスト用フィクスチャキャッシュ層

<!-- 正本(現行仕様)。ここに書かれている内容は人間が承認済みであることが不変条件。
     目的・背景・未決事項・改訂履歴は書かない(それらは docs/spec/changes/ 側の履歴)。
     下のメタ表は features/README.md の索引を機械生成する入力。行の順序と項目名を変えない。 -->

| 項目 | 内容 |
|---|---|
| 対象 | VBA |
| 対象モジュール | TestFixtureCache.bas, TestFixtureCodec.bas |
| 関連SPEC | SPEC-001 |
| 出典 | SPEC-001 |

## 1. 入力

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

## 2. 処理

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
削除する。ローダーの `SCHEMA_VERSION` を上げ忘れた際の復旧手段を兼ねる。

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

## 3. 出力

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

## 4. 異常系・境界値

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

## 5. この機能固有のルール

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
