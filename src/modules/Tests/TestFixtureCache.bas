Attribute VB_Name = "TestFixtureCache"
Option Explicit

'! @brief   テスト用フィクスチャの2層キャッシュ(L1=メモリ / L2=ディスク)
'! @details L1 は Excel プロセス常駐でクラスオブジェクトをそのまま保持する。
'!          xlflow push でモジュールが再インポートされると消えるため、
'!          同じ最終状態を2次元配列にフラット化した L2 をディスクへ持つ。
'!          キャッシュ対象は「外部ブックの生データ」ではなく、突合せ・再設定まで
'!          済ませた最終状態である。詳細は FEAT-001 を参照。
'!
'!          本モジュールはテスト専用。本番モジュールから参照しないこと。

Private Const CACHE_DIR_NAME As String = "test-fixture-cache"

Private Const ERR_SOURCE_MISSING As Long = vbObjectError + 601
Private Const ERR_NOT_2D As Long = vbObjectError + 602

Private mObjects As Object
Private mArrays As Object

'* @brief   キャッシュキーを組み立てる
'* @param   srcPath        展開元の外部ブック(相対パスはリポジトリ直下基準)
'* @param   loaderName     ローダー識別名。キャッシュファイル名の接頭辞になる
'* @param   schemaVersion  加工ロジックの版。ロジックを変えたら呼び出し側が1つ上げる
'* @return  キャッシュキー
'* @details 更新日時とバイト数を含むため元ファイルの変更でキーが変わる。
'*          schemaVersion を含むため、元ファイルが無変更でも版を上げれば変わる。
Public Function BuildKey(ByVal srcPath As String, ByVal loaderName As String, ByVal schemaVersion As Long) As String
    Dim fullPath As String

    fullPath = ResolvePath(srcPath)
    If Not FileExists(fullPath) Then
        Err.Raise ERR_SOURCE_MISSING, "TestFixtureCache.BuildKey", _
            "フィクスチャ元ファイルが見つかりません: " & fullPath
    End If

    BuildKey = loaderName & "|" & LCase$(fullPath) & "|" & _
        Format$(FileDateTime(fullPath), "yyyymmddhhnnss") & "|" & _
        CStr(FileLen(fullPath)) & "|v" & CStr(schemaVersion)
End Function

'* @brief   L1 からオブジェクトを取り出す
'* @param   cacheKey  キャッシュキー
'* @param   outObj    取り出し先。ヒット時のみ設定する
'* @return  ヒットすれば True
'* @details 返るのは L1 が保持している共有参照。呼び出し側が中身を書き換えると
'*          後続テストのフィクスチャが汚染され、順序依存の失敗を招く。
'*          読み取りに限って使うか、書き換えるなら配列から組み直すこと。
Public Function TryGetObject(ByVal cacheKey As String, ByRef outObj As Object) As Boolean
    EnsureStores

    If Not mObjects.Exists(cacheKey) Then
        Exit Function
    End If

    Set outObj = mObjects(cacheKey)
    TryGetObject = True
End Function

'* @brief   L1 へオブジェクトを預ける
'* @param   cacheKey  キャッシュキー
'* @param   obj       預けるオブジェクト
'* @details ディスクへは書かない。クラスのシリアライザを持たない設計のため。
Public Sub PutObject(ByVal cacheKey As String, ByVal obj As Object)
    EnsureStores
    Set mObjects(cacheKey) = obj
End Sub

'* @brief   L1 配列 → L2 ディスク の順に2次元配列を取り出す
'* @param   cacheKey  キャッシュキー
'* @param   outData   取り出し先。ヒット時のみ設定する
'* @return  ヒットすれば True
Public Function TryGetArray(ByVal cacheKey As String, ByRef outData As Variant) As Boolean
    Dim path As String
    Dim text As String
    Dim restored As Variant

    EnsureStores

    If mArrays.Exists(cacheKey) Then
        outData = mArrays(cacheKey)
        TryGetArray = True
        Exit Function
    End If

    path = CacheFilePath(cacheKey)
    If Not FileExists(path) Then
        Exit Function
    End If

    If Not TryReadUtf8(path, text) Then
        DiscardCorrupt path, "読み取りに失敗しました"
        Exit Function
    End If

    If Not TestFixtureCodec.TryDeserialize(text, cacheKey, restored) Then
        DiscardCorrupt path, TestFixtureCodec.LastDiagnostic()
        Exit Function
    End If

    ' 次回以降は L1 で済むよう昇格させる
    mArrays(cacheKey) = restored
    outData = restored
    TryGetArray = True
End Function

'* @brief   2次元配列を L1 と L2 の両方へ格納する
'* @param   cacheKey  キャッシュキー
'* @param   data      2次元配列
'* @details L2 への書き込み失敗はテストを落とさず、記録を残して L1 のみで続行する。
Public Sub PutArray(ByVal cacheKey As String, ByRef data As Variant)
    Dim path As String
    Dim text As String

    If Not TestFixtureCodec.Is2DArray(data) Then
        Err.Raise ERR_NOT_2D, "TestFixtureCache.PutArray", _
            "2次元配列のみキャッシュできます(未初期化配列・1次元配列・スカラーは対象外)。"
    End If

    EnsureStores
    mArrays(cacheKey) = data

    If Not TryEnsureCacheDir() Then
        XlflowDebug.Log "TestFixtureCache: キャッシュディレクトリを作成できません。L1のみで続行します:", CacheDir()
        Exit Sub
    End If

    text = TestFixtureCodec.Serialize(cacheKey, data)
    path = CacheFilePath(cacheKey)
    If Not TryWriteUtf8(path, text) Then
        XlflowDebug.Log "TestFixtureCache: L2への書き込みに失敗しました。L1のみで続行します:", path
    End If
End Sub

'* @brief   L1 を空にする
Public Sub ClearMemory()
    Set mObjects = Nothing
    Set mArrays = Nothing
End Sub

'* @brief   L2 のキャッシュファイルを削除し、L1 も併せて空にする
'* @details ローダーの SCHEMA_VERSION 上げ忘れに気付いたときの復旧手段。
'*          L1 を残すと同一 Excel プロセス内で古い加工結果を返し続け、
'*          復旧手段として成立しないため、両層をまとめて破棄する。
Public Sub ClearDisk()
    Dim dirPath As String
    Dim name As String
    Dim victims As Collection
    Dim victim As Variant

    ClearMemory

    dirPath = CacheDir()
    If Len(FolderName(dirPath)) = 0 Then
        Exit Sub
    End If

    ' Dir$ の列挙中に Kill すると列挙が壊れるため、先に集めてから消す
    Set victims = New Collection
    name = Dir$(dirPath & "\*.txt", vbNormal)
    Do While Len(name) > 0
        victims.Add dirPath & "\" & name
        name = Dir$()
    Loop

    For Each victim In victims
        Kill CStr(victim)
    Next victim
End Sub

'* @brief   L2 キャッシュディレクトリ
'* @return  絶対パス(リポジトリ直下の tmp 配下)
Public Function CacheDir() As String
    CacheDir = RepoRoot() & "\tmp\" & CACHE_DIR_NAME
End Function

'* @brief   L1 の辞書を用意する
Private Sub EnsureStores()
    If mObjects Is Nothing Then
        Set mObjects = CreateObject("Scripting.Dictionary")
    End If
    If mArrays Is Nothing Then
        Set mArrays = CreateObject("Scripting.Dictionary")
    End If
End Sub

'* @brief   キャッシュキーに対応する L2 ファイルのパス
'* @param   cacheKey  キャッシュキー
'* @return  絶対パス
Private Function CacheFilePath(ByVal cacheKey As String) As String
    CacheFilePath = CacheDir() & "\" & SafeLoaderName(cacheKey) & "-" & _
        TestFixtureCodec.KeyHash(cacheKey) & ".txt"
End Function

'* @brief   キャッシュキーの先頭部分をファイル名に使える形へ整える
'* @param   cacheKey  キャッシュキー
'* @return  ファイル名の接頭辞
Private Function SafeLoaderName(ByVal cacheKey As String) As String
    Dim name As String
    Dim separator As Long
    Dim i As Long
    Dim ch As String
    Dim result As String

    name = cacheKey
    separator = InStr(name, "|")
    If separator > 1 Then
        name = Left$(name, separator - 1)
    End If

    ' ファイル名に使えない文字を落とす。取りこぼしの衝突は本文1行目のキー照合で検出する
    For i = 1 To Len(name)
        ch = Mid$(name, i, 1)
        If InStr("\/:*?""<>|", ch) > 0 Then
            result = result & "_"
        Else
            result = result & ch
        End If
    Next i

    If Len(result) = 0 Then
        result = "fixture"
    End If

    SafeLoaderName = result
End Function

'* @brief   破損したキャッシュファイルを記録のうえ破棄する
'* @param   path    対象ファイル
'* @param   reason  破損理由
'* @details 破損はテストを落とさずミス扱いにするが、記録を省くと
'*          「エラーの握りつぶし」になるため Log は必ず出す。
Private Sub DiscardCorrupt(ByVal path As String, ByVal reason As String)
    XlflowDebug.Log "TestFixtureCache: キャッシュ破損のため破棄します:", path, reason

    On Error Resume Next
    Kill path
    Err.Clear
    On Error GoTo 0
End Sub

'* @brief   キャッシュディレクトリを用意する
'* @return  用意できれば True
Private Function TryEnsureCacheDir() As Boolean
    Dim tmpDir As String
    Dim target As String

    tmpDir = RepoRoot() & "\tmp"
    target = CacheDir()

    On Error GoTo Failed
    If Len(FolderName(tmpDir)) = 0 Then
        MkDir tmpDir
    End If
    If Len(FolderName(target)) = 0 Then
        MkDir target
    End If

    TryEnsureCacheDir = True
    Exit Function

Failed:
    Err.Clear
End Function

'* @brief   UTF-8(BOM無し)でテキストを書き出す
'* @param   path  出力先
'* @param   text  内容
'* @return  成功すれば True
Private Function TryWriteUtf8(ByVal path As String, ByVal text As String) As Boolean
    Dim stream As Object
    Dim binary As Object
    Dim errNumber As Long

    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text

    ' ADODB.Stream の utf-8 は BOM を付ける。他ツールでの目視確認を妨げるため取り除く
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3

    Set binary = CreateObject("ADODB.Stream")
    binary.Type = 1
    binary.Open
    stream.CopyTo binary
    binary.SaveToFile path, 2
    binary.Close
    stream.Close

    TryWriteUtf8 = True
    Exit Function

Failed:
    errNumber = Err.Number
    Err.Clear
    CloseStreamQuietly binary
    CloseStreamQuietly stream

    XlflowDebug.Log "TestFixtureCache: UTF-8 書き込みに失敗しました:", path, "err=" & CStr(errNumber)
End Function

'* @brief   ストリームを閉じる。後始末の失敗は元のエラーを覆い隠さないよう握る
'* @param   target  対象ストリーム(Nothing 可)
Private Sub CloseStreamQuietly(ByVal target As Object)
    If target Is Nothing Then
        Exit Sub
    End If

    On Error Resume Next
    target.Close
    On Error GoTo 0
End Sub

'* @brief   UTF-8 のテキストを読み込む
'* @param   path     入力元
'* @param   outText  読み取り結果
'* @return  成功すれば True
Private Function TryReadUtf8(ByVal path As String, ByRef outText As String) As Boolean
    Dim stream As Object

    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    outText = stream.ReadText(-1)
    stream.Close

    TryReadUtf8 = True
    Exit Function

Failed:
    Err.Clear
    CloseStreamQuietly stream
End Function

'* @brief   相対パスをリポジトリ直下基準で絶対化する
'* @param   path  対象パス
'* @return  絶対パス
Private Function ResolvePath(ByVal path As String) As String
    Dim normalized As String

    normalized = Replace$(path, "/", "\")

    If Mid$(normalized, 2, 1) = ":" Or Left$(normalized, 2) = "\\" Then
        ResolvePath = normalized
    Else
        ResolvePath = RepoRoot() & "\" & normalized
    End If
End Function

'* @brief   リポジトリ直下の絶対パス
'* @return  絶対パス(ブックは build\ 配下にある前提)
Private Function RepoRoot() As String
    Dim bookDir As String

    bookDir = ThisWorkbook.Path
    RepoRoot = Left$(bookDir, InStrRev(bookDir, "\") - 1)
End Function

'* @brief   ファイルの存在を確認する
'* @param   path  対象パス
'* @return  存在すれば True
Private Function FileExists(ByVal path As String) As Boolean
    Dim found As String

    On Error GoTo Missing
    found = Dir$(path, vbNormal)
    FileExists = (Len(found) > 0)
    Exit Function

Missing:
    Err.Clear
    FileExists = False
End Function

'* @brief   フォルダ名を取得する(存在しなければ空文字)
'* @param   path  対象パス
'* @return  フォルダ名。無ければ空文字
Private Function FolderName(ByVal path As String) As String
    On Error GoTo Missing
    FolderName = Dir$(path, vbDirectory)
    Exit Function

Missing:
    Err.Clear
    FolderName = ""
End Function
