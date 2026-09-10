Attribute VB_Name = "TestFixtureCacheTests"
Option Explicit

'! @brief   SPEC-001 テスト用フィクスチャキャッシュ層の受け入れテスト
'! @details 外部ブックを実際に生成して読み、実オープン回数を数えることで
'!          「2回目以降は開かない」を直接検証する。

Private Const LOADER_NAME As String = "CacheProbe"
Private Const FIXTURE_ROWS As Long = 1200
Private Const FIXTURE_COLS As Long = 12
Private Const ERR_SOURCE_MISSING As Long = vbObjectError + 601
Private Const ERR_NOT_2D As Long = vbObjectError + 602

Private mFixturePath As String
Private mOpenCount As Long

'* @brief   フィクスチャ元ブックを1度だけ生成する
Public Sub BeforeAll()
    mFixturePath = FixtureDir() & "\fixture-source.xlsx"
    EnsureFixtureWorkbook
End Sub

'* @brief   生成物を残さないよう後始末する
Public Sub AfterAll()
    TestFixtureCache.ClearMemory
    TestFixtureCache.ClearDisk
    DeleteFixtureWorkbook
End Sub

'* @brief   テスト間の独立性を保つため両層を空にする
Public Sub BeforeEach()
    TestFixtureCache.ClearMemory
    TestFixtureCache.ClearDisk
    mOpenCount = 0
End Sub

'* @brief   AC-1: 2回目のロードで実オープンが発生しない
Public Sub Test_Cache_SecondLoad_DoesNotOpenWorkbook()
    Dim first As Variant
    Dim second As Variant

    first = LoadFixture(1)
    XlflowAssert.AssertEquals 1, mOpenCount, "初回は実オープンが1回"

    second = LoadFixture(1)
    XlflowAssert.AssertEquals 1, mOpenCount, "2回目はキャッシュヒットで実オープンが増えない"
    XlflowAssert.AssertEquals first(1, 1), second(1, 1), "同じ内容が返る"
    XlflowAssert.AssertEquals first(FIXTURE_ROWS, FIXTURE_COLS), second(FIXTURE_ROWS, FIXTURE_COLS), "末尾も同じ内容"
End Sub

'* @brief   AC-2: L1 を破棄しても L2 ヒットで実オープンが発生しない
Public Sub Test_Cache_L2Hit_AfterClearMemory()
    Dim first As Variant
    Dim second As Variant
    Dim cacheKey As String
    Dim data As Variant

    first = LoadFixture(1)
    XlflowAssert.AssertEquals 1, mOpenCount, "初回は実オープンが1回"

    ' L1 のみ破棄する。xlflow push でモジュール変数が消える状況を模す
    TestFixtureCache.ClearMemory

    cacheKey = TestFixtureCache.BuildKey(mFixturePath, LOADER_NAME, 1)
    XlflowAssert.AssertTrue TestFixtureCache.TryGetArray(cacheKey, data), "L2 からヒットする"

    TestFixtureCache.ClearMemory
    second = LoadFixture(1)
    XlflowAssert.AssertEquals 1, mOpenCount, "L2 ヒットなので実オープンが増えない"
    XlflowAssert.AssertEquals first(3, 4), second(3, 4), "同じ内容が返る"
End Sub

'* @brief   AC-3: 元ファイルの更新日時が変わるとキャッシュが無効になる
'* @details 更新日時を確実に動かすため約2秒待つ。この分だけ実行が遅い
Public Sub Test_Cache_Invalidated_WhenSourceModified()
    Dim cacheKey As String
    Dim data As Variant

    Call LoadFixture(1)
    cacheKey = TestFixtureCache.BuildKey(mFixturePath, LOADER_NAME, 1)
    XlflowAssert.AssertTrue TestFixtureCache.TryGetArray(cacheKey, data), "前提: キャッシュ済み"

    ' FileDateTime は秒精度のため、確実に更新日時を動かすには1秒以上空ける必要がある
    Application.Wait Now + TimeSerial(0, 0, 2)
    TouchFixtureWorkbook

    XlflowAssert.AssertFalse TestFixtureCache.TryGetArray( _
        TestFixtureCache.BuildKey(mFixturePath, LOADER_NAME, 1), data), _
        "更新日時が変わるとキーが変わりミスになる"
End Sub

'* @brief   AC-4: スキーマ版を上げるとキャッシュが無効になる
Public Sub Test_Cache_Invalidated_WhenSchemaVersionChanged()
    Dim data As Variant

    Call LoadFixture(1)
    XlflowAssert.AssertEquals 1, mOpenCount, "初回は実オープンが1回"

    XlflowAssert.AssertFalse TestFixtureCache.TryGetArray( _
        TestFixtureCache.BuildKey(mFixturePath, LOADER_NAME, 2), data), _
        "スキーマ版が変わると別エントリになる"

    Call LoadFixture(2)
    XlflowAssert.AssertEquals 2, mOpenCount, "版を上げると実ロードが走る"
End Sub

'* @brief   AC-5: L2 往復で型と値が保存される
Public Sub Test_Codec_RoundTrip_PreservesTypes()
    Dim source As Variant
    Dim restored As Variant
    Dim text As String
    Const KEY_NAME As String = "codec|types|v1"

    ReDim source(1 To 1, 1 To 6)
    source(1, 1) = CDbl(1.5)
    source(1, 2) = "abc"
    source(1, 3) = True
    source(1, 4) = Empty
    source(1, 5) = CVErr(2042)
    source(1, 6) = Null

    text = TestFixtureCodec.Serialize(KEY_NAME, source)
    XlflowAssert.AssertTrue TestFixtureCodec.TryDeserialize(text, KEY_NAME, restored), "復元できる"

    XlflowAssert.AssertEquals vbDouble, VarType(restored(1, 1)), "Double の型が保存される"
    XlflowAssert.AssertEquals 1.5, restored(1, 1), "Double の値が保存される"
    XlflowAssert.AssertEquals vbString, VarType(restored(1, 2)), "String の型が保存される"
    XlflowAssert.AssertEquals "abc", restored(1, 2), "String の値が保存される"
    XlflowAssert.AssertEquals vbBoolean, VarType(restored(1, 3)), "Boolean の型が保存される"
    XlflowAssert.AssertEquals True, restored(1, 3), "Boolean の値が保存される"
    XlflowAssert.AssertTrue IsEmpty(restored(1, 4)), "Empty が保存される"
    XlflowAssert.AssertEquals vbError, VarType(restored(1, 5)), "エラー値の型が保存される"
    XlflowAssert.AssertEquals 2042, CLng(restored(1, 5)), "エラー値の番号が保存される"
    XlflowAssert.AssertTrue IsNull(restored(1, 6)), "Null が保存される"
End Sub

'* @brief   有効15桁を超える Double が L2 往復で1ビットも変わらない
'* @details Str$ / CStr は15桁で丸めるため、往復で別の値になりキャッシュヒット時
'*          だけ結果が変わる(偽の緑)。その回帰を防ぐ。
Public Sub Test_Codec_RoundTrip_PreservesFullDoublePrecision()
    Dim source As Variant
    Dim restored As Variant
    Dim text As String
    Dim third As Double
    Dim sum As Double
    Const KEY_NAME As String = "codec|precision|v1"

    third = 1# / 3#
    sum = 0.1 + 0.2

    ReDim source(1 To 1, 1 To 4)
    source(1, 1) = third
    source(1, 2) = sum
    ' 指数リテラルを直接書くと xlflow fmt の operator_spacing が "E+308" の +/- を
    ' 演算子と誤認して "E + 308" に割ってしまい、コードが壊れる。文字列から起こす
    source(1, 3) = Val("1.7976931348623157E+308")
    source(1, 4) = Val("2.2250738585072014E-308")

    text = TestFixtureCodec.Serialize(KEY_NAME, source)
    XlflowAssert.AssertTrue TestFixtureCodec.TryDeserialize(text, KEY_NAME, restored), _
        "復元できる: " & TestFixtureCodec.LastDiagnostic()

    XlflowAssert.AssertEquals third, restored(1, 1), "1/3 が往復で一致する"
    XlflowAssert.AssertEquals sum, restored(1, 2), "0.1+0.2 が往復で一致する"
    XlflowAssert.AssertEquals Val("1.7976931348623157E+308"), restored(1, 3), "最大値が往復で一致する"
    XlflowAssert.AssertEquals Val("2.2250738585072014E-308"), restored(1, 4), "最小正規化数が往復で一致する"
End Sub

'* @brief   破損したヘッダ・本文は例外ではなく False で返る
'* @details 巨大な行数・列数で Long がオーバーフローすると実行時エラー6が
'*          呼び出し側まで飛び、FEAT-001 §4 の「ミス扱い」が破れる。
Public Sub Test_Codec_Rejects_CorruptPayloads()
    Dim restored As Variant
    Dim head As String
    Const KEY_NAME As String = "codec|corrupt|v1"

    head = "XLFCACHE" & vbTab & "1" & vbTab & KEY_NAME & vbLf

    XlflowAssert.AssertFalse TestFixtureCodec.TryDeserialize( _
        head & "99999999999" & vbTab & "99999999999" & vbTab & "1" & vbTab & "1", KEY_NAME, restored), _
        "巨大な行数・列数はミス扱い"

    XlflowAssert.AssertFalse TestFixtureCodec.TryDeserialize( _
        head & "2" & vbTab & "2" & vbTab & "1" & vbTab & "1" & vbLf & "D" & vbTab & "1", KEY_NAME, restored), _
        "レコード数不足はミス扱い"

    XlflowAssert.AssertFalse TestFixtureCodec.TryDeserialize( _
        head & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbLf & "Z" & vbTab & "1", KEY_NAME, restored), _
        "未知の型タグはミス扱い"

    XlflowAssert.AssertFalse TestFixtureCodec.TryDeserialize( _
        head & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbLf & "D" & vbTab & "abc", KEY_NAME, restored), _
        "数値でない D レコードはミス扱い"

    XlflowAssert.AssertFalse TestFixtureCodec.TryDeserialize( _
        head & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbLf & "R" & vbTab & "99999999999", KEY_NAME, restored), _
        "範囲外の R レコードはミス扱い"

    XlflowAssert.AssertNotEqual "", TestFixtureCodec.LastDiagnostic(), "破損理由が記録される"
End Sub

'* @brief   AC-6: 区切り文字とエスケープ文字を含む文字列が往復で壊れない
Public Sub Test_Codec_RoundTrip_EscapesControlChars()
    Dim source As Variant
    Dim restored As Variant
    Dim text As String
    Dim tricky As String
    Const KEY_NAME As String = "codec|escape|v1"

    ' 区切り(TAB/LF)とエスケープ文字そのものが混在しても壊れないことを見る
    tricky = "a" & vbTab & "b" & vbCr & "c" & vbLf & "d\e\\f\tg"

    ReDim source(1 To 1, 1 To 1)
    source(1, 1) = tricky

    text = TestFixtureCodec.Serialize(KEY_NAME, source)
    XlflowAssert.AssertTrue TestFixtureCodec.TryDeserialize(text, KEY_NAME, restored), "復元できる"
    XlflowAssert.AssertEquals tricky, restored(1, 1), "制御文字を含む文字列が一致する"
End Sub

'* @brief   AC-7: 破損した L2 ファイルはミス扱いになり、テストは失敗しない
Public Sub Test_Cache_CorruptFile_FallsBackAndLogs()
    Dim cacheKey As String
    Dim data As Variant
    Dim victim As String

    Call LoadFixture(1)
    cacheKey = TestFixtureCache.BuildKey(mFixturePath, LOADER_NAME, 1)

    victim = FindCacheFile()
    XlflowAssert.AssertNotEqual "", victim, "前提: L2 にキャッシュファイルがある"

    OverwriteWithGarbage victim
    ' L1 を外さないとメモリ側がヒットして L2 の破損まで到達しない
    TestFixtureCache.ClearMemory

    ' 破損は落とさずミス扱い。記録は XlflowDebug.Log 経由で --json の debug 配列に出る
    XlflowAssert.AssertFalse TestFixtureCache.TryGetArray(cacheKey, data), "破損はミス扱いになる"
    XlflowAssert.AssertFalse FileExistsAt(victim), "破損ファイルは破棄される"
End Sub

'* @brief   AC-8: 元ファイル不在で BuildKey がエラーを発生させる
Public Sub Test_Cache_MissingSource_RaisesError()
    Dim errNumber As Long
    Dim missing As String

    missing = FixtureDir() & "\does-not-exist.xlsx"

    On Error Resume Next
    Call TestFixtureCache.BuildKey(missing, LOADER_NAME, 1)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo 0

    XlflowAssert.AssertEquals ERR_SOURCE_MISSING, errNumber, "元ファイル不在は 601"
End Sub

'* @brief   AC-9: 2次元配列でない値は PutArray が拒否する
Public Sub Test_Cache_PutArray_RejectsNon2DValue()
    Dim errNumber As Long
    Dim oneDim As Variant
    Dim unallocated() As Variant
    Dim wrapper As Variant

    oneDim = Array(1, 2, 3)

    On Error Resume Next
    TestFixtureCache.PutArray "probe|1d|v1", oneDim
    errNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    XlflowAssert.AssertEquals ERR_NOT_2D, errNumber, "1次元配列は 602"

    ' 未初期化の動的配列は VBA における「0行」の唯一の表現(SPEC-001 §4 No.10)
    wrapper = unallocated
    On Error Resume Next
    TestFixtureCache.PutArray "probe|empty|v1", wrapper
    errNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    XlflowAssert.AssertEquals ERR_NOT_2D, errNumber, "未初期化配列(0行相当)は 602"

    On Error Resume Next
    TestFixtureCache.PutArray "probe|scalar|v1", 42
    errNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    XlflowAssert.AssertEquals ERR_NOT_2D, errNumber, "スカラーは 602"
End Sub

'* @brief   AC-10: L2 ヒットが実ロードの50%以下の時間で済む
'* @details 実ロードを3回繰り返して平均を取るため、この分だけ実行が遅い
Public Sub Test_Cache_L2Hit_IsFasterThanRealLoad()
    Dim realMs As Double
    Dim hitMs As Double
    Dim started As Double
    Dim data As Variant
    Dim cacheKey As String
    Dim i As Long
    ' Timer の分解能は約15.6ms。1回では粒度が粗すぎるため複数回の平均で判定する
    Const REPEAT As Long = 3

    cacheKey = TestFixtureCache.BuildKey(mFixturePath, LOADER_NAME, 1)

    started = Timer
    For i = 1 To REPEAT
        data = ReadFixtureFromWorkbook()
    Next i
    realMs = (Timer - started) * 1000 / REPEAT

    TestFixtureCache.PutArray cacheKey, data

    started = Timer
    For i = 1 To REPEAT
        ' L1 を外して L2 単体の実力を測る
        TestFixtureCache.ClearMemory
        XlflowAssert.AssertTrue TestFixtureCache.TryGetArray(cacheKey, data), "L2 ヒットする"
    Next i
    hitMs = (Timer - started) * 1000 / REPEAT

    ' 測定が空振り(早期 Exit で中身を作っていない)でないことを確かめる
    XlflowAssert.AssertEquals FIXTURE_ROWS, UBound(data, 1), "L2 から全行が復元されている"
    XlflowAssert.AssertEquals FIXTURE_COLS, UBound(data, 2), "L2 から全列が復元されている"
    XlflowAssert.AssertEquals CDbl(FIXTURE_ROWS * 100 + FIXTURE_COLS), data(FIXTURE_ROWS, FIXTURE_COLS), "末尾セルの値が一致する"

    XlflowDebug.Log "AC-10 L2 hit:", Format$(hitMs, "0.0") & "ms", "real load:", Format$(realMs, "0.0") & "ms"
    XlflowAssert.AssertTrue hitMs <= realMs * 0.5, _
        "L2ヒット " & Format$(hitMs, "0.0") & "ms が実ロード " & Format$(realMs, "0.0") & "ms の50%以下であること"
End Sub

'* @brief   AC-11: キャッシュ出力先が tmp 配下である
Public Sub Test_Cache_CacheDir_IsUnderTmp()
    Dim expected As String

    expected = RepoRoot() & "\tmp\test-fixture-cache"
    XlflowAssert.AssertEquals LCase$(expected), LCase$(TestFixtureCache.CacheDir()), _
        "キャッシュ出力先はリポジトリ直下の tmp 配下"
End Sub

'* @brief   L1 にクラスオブジェクトを預けて取り出せる
Public Sub Test_Cache_ObjectLayer_RoundTrips()
    Dim stored As Object
    Dim fetched As Object
    Const KEY_NAME As String = "probe|object|v1"

    Set stored = CreateObject("Scripting.Dictionary")
    stored.Add "answer", 42

    XlflowAssert.AssertFalse TestFixtureCache.TryGetObject(KEY_NAME, fetched), "未登録ならミス"

    TestFixtureCache.PutObject KEY_NAME, stored
    XlflowAssert.AssertTrue TestFixtureCache.TryGetObject(KEY_NAME, fetched), "登録後はヒット"
    XlflowAssert.AssertEquals 42, fetched("answer"), "同じ内容が返る"

    TestFixtureCache.ClearMemory
    XlflowAssert.AssertFalse TestFixtureCache.TryGetObject(KEY_NAME, fetched), "ClearMemory で消える"
End Sub

'* @brief   キャッシュ経由でフィクスチャを読む。ミス時のみ外部ブックを開く
'* @param   schemaVersion  加工ロジックの版
'* @return  フィクスチャの2次元配列
Private Function LoadFixture(ByVal schemaVersion As Long) As Variant
    Dim cacheKey As String
    Dim data As Variant

    cacheKey = TestFixtureCache.BuildKey(mFixturePath, LOADER_NAME, schemaVersion)
    If TestFixtureCache.TryGetArray(cacheKey, data) Then
        LoadFixture = data
        Exit Function
    End If

    data = ReadFixtureFromWorkbook()
    TestFixtureCache.PutArray cacheKey, data
    LoadFixture = data
End Function

'* @brief   外部ブックを実際に開いて一括読みする。実オープン回数を数える
'* @return  読み取った2次元配列
Private Function ReadFixtureFromWorkbook() As Variant
    Dim wb As Object
    Dim prevCalc As Long
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim result As Variant
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating

    On Error GoTo Cleanup
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    mOpenCount = mOpenCount + 1
    Set wb = Application.Workbooks.Open(mFixturePath, 0, True)
    ' 逐次アクセスは桁違いに遅いため一括で取得する
    result = wb.Worksheets(1).Range("A1").Resize(FIXTURE_ROWS, FIXTURE_COLS).Value2

Cleanup:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close False
        Set wb = Nothing
        Err.Clear
        On Error GoTo 0
    End If

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If

    ReadFixtureFromWorkbook = result
End Function

'* @brief   フィクスチャ元ブックを生成する(既にあれば何もしない)
Private Sub EnsureFixtureWorkbook()
    Dim wb As Object
    Dim values As Variant
    Dim r As Long
    Dim c As Long
    Dim prevAlerts As Boolean

    If FileExistsAt(mFixturePath) Then
        Exit Sub
    End If
    EnsureFolder FixtureDir()

    ReDim values(1 To FIXTURE_ROWS, 1 To FIXTURE_COLS)
    For r = 1 To FIXTURE_ROWS
        For c = 1 To FIXTURE_COLS
            values(r, c) = CDbl(r * 100 + c)
        Next c
    Next r
    ' 文字列も混ぜないと型ごとの往復が実データで確認できない
    values(1, 1) = "見出し"

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Range("A1").Resize(FIXTURE_ROWS, FIXTURE_COLS).Value2 = values
    wb.SaveAs mFixturePath, 51
    wb.Close False
    Set wb = Nothing
    Application.DisplayAlerts = prevAlerts
End Sub

'* @brief   内容を保ったまま更新日時だけを現在時刻へ動かす
'* @details 先頭1バイトを読んで同じ値を書き戻す。バイト数も内容も変わらないため、
'*          更新日時だけによる無効化を切り分けて検証できる。
'*          FileCopy は使えない(Windows の CopyFile は更新日時を引き継ぐため)。
Private Sub TouchFixtureWorkbook()
    Dim handle As Integer
    Dim firstByte As Byte

    handle = FreeFile
    Open mFixturePath For Binary Access Read Write As #handle
    Get #handle, 1, firstByte
    Put #handle, 1, firstByte
    Close #handle
End Sub

'* @brief   フィクスチャ元ブックを削除する
Private Sub DeleteFixtureWorkbook()
    If FileExistsAt(mFixturePath) Then
        Kill mFixturePath
    End If
End Sub

'* @brief   L2 キャッシュディレクトリから本テストのキャッシュファイルを1つ探す
'* @return  絶対パス。見つからなければ空文字
Private Function FindCacheFile() As String
    Dim dirPath As String
    Dim name As String

    dirPath = TestFixtureCache.CacheDir()
    name = Dir$(dirPath & "\" & LOADER_NAME & "-*.txt", vbNormal)
    If Len(name) = 0 Then
        Exit Function
    End If

    FindCacheFile = dirPath & "\" & name
End Function

'* @brief   キャッシュファイルの中身を壊す
'* @param   path  対象ファイル
Private Sub OverwriteWithGarbage(ByVal path As String)
    Dim handle As Integer

    handle = FreeFile
    Open path For Output As #handle
    Print #handle, "GARBAGE"
    Close #handle
End Sub

'* @brief   フィクスチャ元ブックの置き場
'* @return  絶対パス
Private Function FixtureDir() As String
    FixtureDir = RepoRoot() & "\tmp\test-fixture-src"
End Function

'* @brief   リポジトリ直下の絶対パス(ブックは build\ 配下にある前提)
'* @return  絶対パス
Private Function RepoRoot() As String
    Dim bookDir As String

    bookDir = ThisWorkbook.Path
    RepoRoot = Left$(bookDir, InStrRev(bookDir, "\") - 1)
End Function

'* @brief   ファイルの存在を確認する
'* @param   path  対象パス
'* @return  存在すれば True
Private Function FileExistsAt(ByVal path As String) As Boolean
    Dim found As String

    On Error GoTo Missing
    found = Dir$(path, vbNormal)
    FileExistsAt = (Len(found) > 0)
    Exit Function

Missing:
    Err.Clear
    FileExistsAt = False
End Function

'* @brief   フォルダを作成する(既にあれば何もしない)
'* @param   path  作成するフォルダ
Private Sub EnsureFolder(ByVal path As String)
    Dim parent As String

    If Len(Dir$(path, vbDirectory)) > 0 Then
        Exit Sub
    End If

    parent = Left$(path, InStrRev(path, "\") - 1)
    If Len(Dir$(parent, vbDirectory)) = 0 Then
        MkDir parent
    End If
    MkDir path
End Sub
