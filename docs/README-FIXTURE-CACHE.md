# テスト用フィクスチャキャッシュ 導入手順

外部Excelブックを開いて内部データへ展開する処理が遅く、テストの反復速度を律速している環境へ、
`TestFixtureCache` 一式を持ち込んで使えるようにするまでの手順。

仕様の正本は [docs/spec/features/FEAT-001-test-fixture-cache.md](spec/features/FEAT-001-test-fixture-cache.md)。
API の引数・戻り値・異常系はそちらが唯一の真実源であり、本書は**移植と運用**だけを扱う。

---

## 手順1: 前提を満たす

| 項目 | 要件 | 確認方法 |
|---|---|---|
| VBA | 3ファイル本体は VBA6 でも動く（`PtrSafe` 宣言を持たない）。`XlflowDebug` をそのまま使う場合のみ VBA 7.1 以降 | `Application.Version` が 15.0 以上なら VBA 7.1 |
| `Scripting.Dictionary` | 遅延バインドで生成できる | `CreateObject("Scripting.Dictionary")` が成功する |
| `ADODB.Stream` | 遅延バインドで生成できる | `CreateObject("ADODB.Stream")` が成功する |
| 書き込み先 | キャッシュ置き場に書き込める | 手順3で決める |

どちらも `CreateObject` の遅延バインドで使うため、**VBE の参照設定は不要**。
ADODB.Stream は Windows 標準構成要素で、ACE OLEDB のような別途インストールは要らない。

完了条件: 上記4項目すべてを確認した。

## 手順2: 3ファイルを持ち込む

| ファイル | 役割 |
|---|---|
| [src/modules/Tests/TestFixtureCache.bas](../src/modules/Tests/TestFixtureCache.bas) | L1/L2 の格納・取得、キー生成、無効化、破棄 |
| [src/modules/Tests/TestFixtureCodec.bas](../src/modules/Tests/TestFixtureCodec.bas) | 2次元配列 ⇔ テキストの直列化・復元 |
| [src/modules/Tests/TestFixtureCacheTests.bas](../src/modules/Tests/TestFixtureCacheTests.bas) | 自己テスト14本。手順4で使う |

依存はこの3本と `XlflowDebug.Log` / `XlflowAssert` のみ。業務モジュールへの依存は無い。

**xlflow を使わない環境へ移す場合**は2箇所を置き換える。

- `XlflowDebug.Log` → その環境のログ出力（`Debug.Print` など）
- `XlflowAssert.*` → その環境のアサーション（自己テストを使う場合のみ）

完了条件: 3ファイルをインポートし、コンパイルが通った。

## 手順3: キャッシュ出力先を環境に合わせる

**移植でいちばん壊れやすいのはここ。** `TestFixtureCache.RepoRoot` は
`ThisWorkbook.Path` の**親ディレクトリ**を返す。この repo はブックが `build/` 配下にあるため
親がリポジトリ直下になるが、ブックの置き場所が違う環境ではキャッシュが意図しない場所に出る。

```vb
' 現状の実装(この repo 前提)
Private Function RepoRoot() As String
    Dim bookDir As String

    bookDir = ThisWorkbook.Path
    RepoRoot = Left$(bookDir, InStrRev(bookDir, "\") - 1)
End Function
```

移植先での置き換え例。

```vb
' ブックと同じ階層に tmp\ を作る場合
Private Function RepoRoot() As String
    RepoRoot = ThisWorkbook.Path
End Function

' 置き場所を固定する場合
Private Function RepoRoot() As String
    RepoRoot = "D:\work\fixture-cache-root"
End Function
```

決めたら**バージョン管理の対象外にする**。この repo は `.gitignore` の `tmp/` が効くため、
`CacheDir` が `tmp\test-fixture-cache` を指すことで混入経路が構造的に消えている。
移植先でも同じ性質を作る（`.gitignore` に加える、リポジトリ外に置く、のいずれか）。

完了条件: `TestFixtureCache.CacheDir()` が返すパスを実際に確認し、
そのパスがバージョン管理の対象外であることを確かめた。

## 手順4: 自己テストで動作確認する

`TestFixtureCacheTests` は**外部ブックを自前で生成して**検証するため、実際のマスタブックが
無くても動く。移植直後の健全性確認に使う。

```bash
xlflow test --module TestFixtureCacheTests --session --json
```

xlflow を使わない環境では、`BeforeAll` → 各 `Test_*` → `AfterAll` の順に手で実行する。

14本すべて PASS すれば、その環境で L1/L2・無効化・破損フォールバック・型の往復が
成立している。所要時間の実測値も出る。

```
AC-10 L2 hit: 7.8ms real load: 187.5ms
```

完了条件: 14/14 PASS。1本でも落ちたら手順3のパス設定か手順1の前提を疑う。

## 手順5: ローダーを1本書く

フィクスチャ1つにつき標準モジュール1本。`SRC_PATH` と2つの `Build*` を埋める。

```vb
Attribute VB_Name = "TestFixtureOrders"
Option Explicit

'! @brief 受注マスタのテスト用フィクスチャ

'* 加工ロジック(突合せ・再設定)を変更したら +1 する。
'* 据え置くと外部ブック無変更のまま古い加工結果がヒットし、テストが偽の緑になる
Private Const SCHEMA_VERSION As Long = 1
Private Const LOADER_NAME As String = "TestFixtureOrders"
Private Const SRC_PATH As String = "tests\fixtures\orders.xlsx"

'* @brief   受注マスタを返す。2回目以降は外部ブックを開かない
'* @return  受注番号をキーにした Dictionary
Public Function Orders() As Object
    Dim cacheKey As String
    Dim cached As Object
    Dim data As Variant

    cacheKey = TestFixtureCache.BuildKey(SRC_PATH, LOADER_NAME, SCHEMA_VERSION)

    If TestFixtureCache.TryGetObject(cacheKey, cached) Then
        Set Orders = cached
        Exit Function
    End If

    If Not TestFixtureCache.TryGetArray(cacheKey, data) Then
        data = BuildFromWorkbook()
        TestFixtureCache.PutArray cacheKey, data
    End If

    Set cached = BuildObjects(data)
    TestFixtureCache.PutObject cacheKey, cached
    Set Orders = cached
End Function
```

**`BuildFromWorkbook` に、外部ブックを開く処理と突合せ・再設定までを全部入れる。**
ここが返す2次元配列がそのまま L2 に載るため、加工コストごと省ける。
生データを返して呼び出し側で加工すると、加工コストは毎回かかる。

`BuildFromWorkbook` の Excel 状態の抑止と COM 解放は、
[TestFixtureCacheTests.bas](../src/modules/Tests/TestFixtureCacheTests.bas) の
`ReadFixtureFromWorkbook` をひな型にする（`Cleanup:` ラベルで元の値へ必ず戻す形）。

完了条件: `Orders()` を2回呼び、2回目に `Workbooks.Open` が発生しないことを確認した。

## 手順6: テストから使う

```vb
Public Sub Test_月締め_赤伝を除外する()
    Dim orders As Object

    Set orders = TestFixtureOrders.Orders()
    ' 以降、orders を読んで検証する
End Sub
```

業務テストの `BeforeEach` では**両層をそのまま残す**。テスト間でキャッシュが生き続けることが
速度の源になる。`TestFixtureCacheTests` が `BeforeEach` で毎回消しているのは、
キャッシュ機構そのものを検証する目的のためであり、業務テストの手本ではない。

`TryGetObject` が返すのは**共有参照**。返った `Dictionary` やクラスを書き換えると後続テストの
フィクスチャが汚染され、順序依存の偽の緑になる。読むだけで使うか、書き換える必要があるなら
`BuildObjects` で毎回組み直す。

完了条件: 同じフィクスチャを使うテストを2本以上動かし、実オープンが1回に収まった。

---

## 運用

覚えることは1つ。**加工ロジックを変えたら `SCHEMA_VERSION` を +1 する。**

外部ブックの更新は `FileDateTime` と `FileLen` がキーに入っているため自動で検出される。
手作業は要らない。

キャッシュを疑ったときの復旧はこれ。L1 も消さないと同一 Excel プロセス内では
古い結果が返り続ける。

```vb
TestFixtureCache.ClearMemory
TestFixtureCache.ClearDisk
```

## 経路と体感速度

| 状況 | 経路 | 実測の目安(1200行×12列) |
|---|---|---|
| 初回 | 外部ブックを開く | 187ms |
| 同一テスト実行内の2回目以降 | L1(メモリ) | ほぼ0ms |
| モジュール再インポート後 | L2(ディスク) | 7.8ms |
| 外部ブック更新後 | 自動で無効化 → 再ロード | 187ms |

## トラブルシューティング

| 症状 | 原因 | 対処 |
|---|---|---|
| 毎回 `Workbooks.Open` が走る | `SRC_PATH` の解決先が呼び出しごとに違う | `BuildKey` に渡す前に絶対パス化する |
| キャッシュが想定外の場所に出る | `RepoRoot` がブック配置と合っていない | 手順3で置き換える |
| 外部ブックを直したのに古い結果 | 1秒以内かつ同一バイト数の更新 | `FileDateTime` が秒精度のため検出できない。1秒以上空けて保存し直す |
| ロジックを直したのに古い結果 | `SCHEMA_VERSION` が据え置き | +1 する。応急処置は `ClearMemory` + `ClearDisk` |
| テストが順序で結果が変わる | `TryGetObject` の共有参照を書き換えている | 読み取り専用で使うか毎回組み直す |
| `キャッシュ破損のため破棄します` がログに出続ける | `LOADER_NAME` に TAB / 改行 / `\|` が入っている | 英数字とハイフンに収める |
| L2 だけ効かない | キャッシュ置き場に書き込めていない | ログの `L2への書き込みに失敗しました` を確認し、書き込み権限を見る |

## 既知の制約

- **1秒以内かつ同一バイト数の更新は検出できない**（`FileDateTime` が秒精度のため）。
- **`SCHEMA_VERSION` の上げ忘れは機械検出できない**。規約とレビューで担保する。
  保険は `ClearDisk`。
- **キャッシュファイルは溜まり続ける**。世代管理は持たないため、増えたら `ClearDisk` を叩く。
- **`xlflow fmt` は指数リテラルを壊す**。`operator_spacing` が `1.5E+308` の `+` を
  演算子と誤認して `E + 308` に割る。VBA ソースには `Val("1.5E+308")` の形で書く。
- **`'@Tag("...")` と Doxygen ヘッダは併用できない**。`scripts/check_doxygen.py` は宣言の
  直前行が `'*` であることを求め、xlflow の Tag 検出は直前行が `'@` であることを求めるため、
  同じ行を奪い合う。この repo は Doxygen を優先している。
