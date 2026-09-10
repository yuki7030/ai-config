Attribute VB_Name = "TestFixtureCodec"
Option Explicit

'! @brief   TestFixtureCache の L2 永続に使う 2次元配列 ⇔ テキスト変換
'! @details 書式は FEAT-001「3. 出力」を参照。型タグ付きの行指向テキストで、
'!          Range.Value2 が返しうる Double / String / Boolean / Empty / Null /
'!          エラー値の6種だけを扱う。

Private Const CACHE_MAGIC As String = "XLFCACHE"
Private Const CACHE_FORMAT_VERSION As String = "1"

Private Const ERR_NOT_2D As Long = vbObjectError + 602
Private Const ERR_UNSUPPORTED_TYPE As Long = vbObjectError + 603

' VBA の Str$ / CStr / Format$ はいずれも有効15桁までしか出せない。
' 15桁で往復できない Double はビット表現(16進16文字)へ退避するため、
' 同じサイズの2つの型を LSet で重ねて8バイトを取り出す
Private Type TDoubleValue
    value As Double
End Type

Private Type TDoubleBytes
    bytes(0 To 7) As Byte
End Type

' 破損ファイルの巨大値で Long がオーバーフローしないための上限
Private Const MAX_DIMENSION As Double = 1048576#
Private Const MAX_CELLS As Double = 20000000#
Private Const MAX_ERROR_CODE As Double = 65535#

' 型タグの文字コード。復元ループは全セル分回るため文字列比較を避ける
Private Const TAG_DOUBLE As Long = 68
Private Const TAG_STRING As Long = 83
Private Const TAG_BOOLEAN As Long = 66
Private Const TAG_EMPTY As Long = 69
Private Const TAG_NULL As Long = 78
Private Const TAG_ERROR As Long = 82
Private Const TAG_HEX As Long = 72

' ビット表現は 8バイト = 16進16文字ちょうど
Private Const HEX_DOUBLE_LENGTH As Long = 16

Private mLastDiagnostic As String

'* @brief   2次元配列を L2 永続用のテキストへ直列化する
'* @param   cacheKey  ヘッダに埋め込むキャッシュキー
'* @param   data      2次元配列
'* @return  vbLf 区切りのテキスト
Public Function Serialize(ByVal cacheKey As String, ByRef data As Variant) As String
    Dim rowCount As Long
    Dim colCount As Long
    Dim lbRow As Long
    Dim lbCol As Long
    Dim r As Long
    Dim c As Long
    Dim parts() As String
    Dim index As Long

    If Not Is2DArray(data) Then
        Err.Raise ERR_NOT_2D, "TestFixtureCodec.Serialize", _
            "2次元配列のみ直列化できます(未初期化配列・1次元配列・スカラーは対象外)。"
    End If

    lbRow = LBound(data, 1)
    lbCol = LBound(data, 2)
    rowCount = UBound(data, 1) - lbRow + 1
    colCount = UBound(data, 2) - lbCol + 1

    ' セル数が多いと & 連結は O(n^2) になるため、配列へ貯めて Join する
    ReDim parts(0 To rowCount * colCount + 1)
    parts(0) = CACHE_MAGIC & vbTab & CACHE_FORMAT_VERSION & vbTab & cacheKey
    parts(1) = CStr(rowCount) & vbTab & CStr(colCount) & vbTab & CStr(lbRow) & vbTab & CStr(lbCol)

    index = 2
    For r = lbRow To UBound(data, 1)
        For c = lbCol To UBound(data, 2)
            parts(index) = EncodeCell(data(r, c))
            index = index + 1
        Next c
    Next r

    Serialize = Join(parts, vbLf)
End Function

'* @brief   直列化テキストを2次元配列へ復元する
'* @param   text         直列化テキスト
'* @param   expectedKey  期待するキャッシュキー
'* @param   outData      復元先。成功時のみ設定する
'* @return  復元できれば True。失敗理由は LastDiagnostic で取れる
Public Function TryDeserialize(ByVal text As String, ByVal expectedKey As String, ByRef outData As Variant) As Boolean
    Dim lines() As String
    Dim header() As String
    Dim dims() As String
    Dim rowCount As Long
    Dim colCount As Long
    Dim lbRow As Long
    Dim lbCol As Long
    Dim arr As Variant
    Dim r As Long
    Dim c As Long
    Dim index As Long
    Dim line As String
    Dim raw As String
    Dim rowValue As Double
    Dim colValue As Double
    Dim lbRowValue As Double
    Dim lbColValue As Double
    Dim errValue As Double
    Dim doubleValue As Double

    mLastDiagnostic = ""
    lines = Split(text, vbLf)

    If UBound(lines) < 1 Then
        mLastDiagnostic = "ヘッダ行が不足しています"
        Exit Function
    End If

    header = Split(lines(0), vbTab)
    If UBound(header) <> 2 Then
        mLastDiagnostic = "ヘッダ行の項目数が不正です"
        Exit Function
    End If
    If header(0) <> CACHE_MAGIC Then
        mLastDiagnostic = "識別子が一致しません"
        Exit Function
    End If
    If header(1) <> CACHE_FORMAT_VERSION Then
        mLastDiagnostic = "書式版が一致しません: " & header(1)
        Exit Function
    End If
    If header(2) <> expectedKey Then
        mLastDiagnostic = "キャッシュキーが一致しません(ハッシュ衝突)"
        Exit Function
    End If

    dims = Split(lines(1), vbTab)
    If UBound(dims) <> 3 Then
        mLastDiagnostic = "次元行の項目数が不正です"
        Exit Function
    End If

    ' 破損ファイルの巨大値で CLng がオーバーフロー(実行時エラー6)し、
    ' 「破損はミス扱い」の経路を突き破って呼び出し側へ伝播するのを防ぐ。
    ' Val は Double を返すので、Long へ落とす前に範囲を検査する
    rowValue = Val(dims(0))
    colValue = Val(dims(1))
    lbRowValue = Val(dims(2))
    lbColValue = Val(dims(3))

    ' 0行の2次元配列は VBA に存在しないため、保存側でも弾いている(SPEC-001 §4 No.10)
    If rowValue < 1 Or colValue < 1 Or rowValue > MAX_DIMENSION Or colValue > MAX_DIMENSION Then
        mLastDiagnostic = "行数・列数が不正です"
        Exit Function
    End If
    If rowValue * colValue > MAX_CELLS Then
        mLastDiagnostic = "セル数が上限(" & CStr(MAX_CELLS) & ")を超えています"
        Exit Function
    End If
    If Abs(lbRowValue) > MAX_DIMENSION Or Abs(lbColValue) > MAX_DIMENSION Then
        mLastDiagnostic = "配列の下限が不正です"
        Exit Function
    End If

    rowCount = CLng(rowValue)
    colCount = CLng(colValue)
    lbRow = CLng(lbRowValue)
    lbCol = CLng(lbColValue)

    If UBound(lines) - 1 <> rowCount * colCount Then
        mLastDiagnostic = "レコード数が宣言(" & CStr(rowCount * colCount) & ")と一致しません: " & CStr(UBound(lines) - 1)
        Exit Function
    End If

    ReDim arr(lbRow To lbRow + rowCount - 1, lbCol To lbCol + colCount - 1)
    index = 2
    For r = lbRow To lbRow + rowCount - 1
        For c = lbCol To lbCol + colCount - 1
            line = lines(index)
            ' 型タグは1文字固定。セル数だけ繰り返す内側なので InStr と関数呼び出しを避け、
            ' 位置決め打ち + 文字コード比較で復元する
            If Mid$(line, 2, 1) <> vbTab Then
                mLastDiagnostic = "レコード書式が不正です(行 " & CStr(index + 1) & ")"
                Exit Function
            End If

            raw = Mid$(line, 3)
            Select Case AscW(line)
            Case TAG_DOUBLE
                ' Val は解釈できない文字列に無言で 0 を返す。破損値が 0 として
                ' L1 へ昇格し、キャッシュ全体へ広がるのを防ぐため書式を検査する
                If Not IsNumeric(raw) Then
                    mLastDiagnostic = "数値として解釈できません(行 " & CStr(index + 1) & "): " & raw
                    Exit Function
                End If
                ' Val は常に "." を小数点として解釈する。書き出し側と対になる
                arr(r, c) = Val(raw)
            Case TAG_HEX
                If Not TryHexToDouble(raw, doubleValue) Then
                    mLastDiagnostic = "ビット表現が不正です(行 " & CStr(index + 1) & "): " & raw
                    Exit Function
                End If
                arr(r, c) = doubleValue
            Case TAG_STRING
                arr(r, c) = Unescape(raw)
            Case TAG_BOOLEAN
                If raw = "1" Then
                    arr(r, c) = True
                ElseIf raw = "0" Then
                    arr(r, c) = False
                Else
                    mLastDiagnostic = "真偽値が不正です(行 " & CStr(index + 1) & "): " & raw
                    Exit Function
                End If
            Case TAG_EMPTY
                arr(r, c) = Empty
            Case TAG_NULL
                arr(r, c) = Null
            Case TAG_ERROR
                errValue = Val(raw)
                ' CVErr の有効範囲外だと実行時エラー5になり、破損時のミス扱いを突き破る
                If Not IsNumeric(raw) Or errValue < 0 Or errValue > MAX_ERROR_CODE Then
                    mLastDiagnostic = "エラー値が範囲外です(行 " & CStr(index + 1) & "): " & raw
                    Exit Function
                End If
                arr(r, c) = CVErr(CLng(errValue))
            Case Else
                mLastDiagnostic = "型タグが不正です(行 " & CStr(index + 1) & "): " & Left$(line, 1)
                Exit Function
            End Select

            index = index + 1
        Next c
    Next r

    outData = arr
    TryDeserialize = True
End Function

'* @brief   値が要素を持つ2次元配列かを判定する
'* @param   value  判定対象
'* @return  2次元配列なら True
Public Function Is2DArray(ByRef value As Variant) As Boolean
    If IsObject(value) Then
        Exit Function
    End If
    If Not IsArray(value) Then
        Exit Function
    End If

    ' 未初期化の動的配列と1次元配列は UBound(,2) がエラー9になる。
    ' VBA に割り当て状態を直接調べる手段が無いため、エラーの有無で判定する
    On Error Resume Next
    Is2DArray = (UBound(value, 2) >= LBound(value, 2))
    Err.Clear
    On Error GoTo 0
End Function

'* @brief   キャッシュキーから16進8桁のハッシュを得る(FNV-1a 32bit)
'* @param   value  ハッシュ対象
'* @return  16進8桁。衝突はキャッシュ本文1行目のキー照合で検出する
Public Function KeyHash(ByVal value As String) As String
    Dim hashHigh As Long
    Dim hashLow As Long
    Dim i As Long
    Dim code As Long

    ' FNV-1a のオフセット基底 0x811C9DC5。VBA の Long は符号付き32bitで
    ' 乗算が容易に溢れるため、上位/下位16bitに分けて保持する
    hashHigh = &H811C&
    hashLow = &H9DC5&

    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1)) And &HFFFF&
        ' UTF-16LE のバイト順に合わせて下位バイトから食わせる
        Fnv1aStep hashHigh, hashLow, code And &HFF&
        Fnv1aStep hashHigh, hashLow, (code \ &H100&) And &HFF&
    Next i

    KeyHash = Right$("0000" & Hex$(hashHigh), 4) & Right$("0000" & Hex$(hashLow), 4)
End Function

'* @brief   直近の復元失敗の理由
'* @return  理由。無ければ空文字
Public Function LastDiagnostic() As String
    LastDiagnostic = mLastDiagnostic
End Function

'* @brief   FNV-1a の1バイト分を進める
'* @param   hashHigh   ハッシュ上位16bit(更新される)
'* @param   hashLow    ハッシュ下位16bit(更新される)
'* @param   byteValue  取り込むバイト
Private Sub Fnv1aStep(ByRef hashHigh As Long, ByRef hashLow As Long, ByVal byteValue As Long)
    Const PRIME_LOW As Long = &H193&
    Const PRIME_HIGH As Long = &H100&
    Dim product As Long
    Dim newLow As Long
    Dim newHigh As Long

    hashLow = hashLow Xor byteValue
    product = hashLow * PRIME_LOW
    newLow = product And &HFFFF&
    ' 16777619 = 0x01000193。上位は 桁上がり + 交差項 で求まる
    newHigh = (hashHigh * PRIME_LOW + hashLow * PRIME_HIGH + (product \ &H10000)) And &HFFFF&

    hashLow = newLow
    hashHigh = newHigh
End Sub

'* @brief   1セルを型タグ付きレコードへ変換する
'* @param   value  セルの値
'* @return  "<型タグ><TAB><値>"
Private Function EncodeCell(ByRef value As Variant) As String
    Dim decimalText As String

    If IsObject(value) Then
        Err.Raise ERR_UNSUPPORTED_TYPE, "TestFixtureCodec.EncodeCell", "オブジェクトは直列化できません。"
    End If

    If IsError(value) Then
        EncodeCell = "R" & vbTab & CStr(CLng(value))
        Exit Function
    End If
    If IsNull(value) Then
        EncodeCell = "N" & vbTab
        Exit Function
    End If
    If IsEmpty(value) Then
        EncodeCell = "E" & vbTab
        Exit Function
    End If

    Select Case VarType(value)
    Case vbString
        EncodeCell = "S" & vbTab & Escape(CStr(value))
    Case vbBoolean
        If CBool(value) Then
            EncodeCell = "B" & vbTab & "1"
        Else
            EncodeCell = "B" & vbTab & "0"
        End If
    Case vbDouble
        ' Str$ はロケールに依らず "." を小数点に使う。CStr は地域設定に従うため避ける
        decimalText = Trim$(Str$(value))
        If DecimalRoundTrips(decimalText, value) Then
            ' 金額・数量・日付シリアルなど実データの大半はここを通り、目視で読める
            EncodeCell = "D" & vbTab & decimalText
        Else
            ' 15桁では往復しない値(0.1+0.2 や 1/3 など)だけビット表現へ退避する。
            ' 10進のまま書くとキャッシュヒット時だけ値が変わる(偽の緑)ため
            EncodeCell = "H" & vbTab & DoubleToHex(value)
        End If
    Case Else
        Err.Raise ERR_UNSUPPORTED_TYPE, "TestFixtureCodec.EncodeCell", _
            "未対応の型です(VarType=" & CStr(VarType(value)) & ")。Range.Value2 由来の値のみ扱えます。"
    End Select
End Function

'* @brief   10進表記が元の Double へ正確に戻るかを判定する
'* @param   text      10進表記
'* @param   expected  元の値
'* @return  完全に一致すれば True
'* @details Double 最大値の近傍では 15桁への丸めが表現可能範囲を超え、
'*          Val がオーバーフローする。その場合もビット表現へ回すため False を返す。
Private Function DecimalRoundTrips(ByVal text As String, ByVal expected As Double) As Boolean
    On Error GoTo Invalid
    DecimalRoundTrips = (Val(text) = expected)
    Exit Function

Invalid:
    Err.Clear
End Function

'* @brief   Double を16進16文字のビット表現へ変換する
'* @param   value  対象の値
'* @return  16進16文字
Private Function DoubleToHex(ByVal value As Double) As String
    Dim source As TDoubleValue
    Dim target As TDoubleBytes
    Dim parts(0 To 7) As String
    Dim i As Long

    source.value = value
    LSet target = source

    For i = 0 To 7
        parts(i) = Right$("0" & Hex$(target.bytes(i)), 2)
    Next i

    DoubleToHex = Join(parts, "")
End Function

'* @brief   16進16文字のビット表現を Double へ戻す
'* @param   text      16進16文字
'* @param   outValue  復元先。成功時のみ設定する
'* @return  復元できれば True
Private Function TryHexToDouble(ByVal text As String, ByRef outValue As Double) As Boolean
    Dim source As TDoubleBytes
    Dim target As TDoubleValue
    Dim i As Long

    If Len(text) <> HEX_DOUBLE_LENGTH Then
        Exit Function
    End If

    On Error GoTo Invalid
    For i = 0 To 7
        source.bytes(i) = CByte(CLng("&H" & Mid$(text, i * 2 + 1, 2)))
    Next i
    LSet target = source

    outValue = target.value
    TryHexToDouble = True
    Exit Function

Invalid:
    Err.Clear
End Function

'* @brief   区切り文字とエスケープ文字を退避する
'* @param   value  元の文字列
'* @return  エスケープ済み文字列
Private Function Escape(ByVal value As String) As String
    value = Replace$(value, "\", "\\")
    value = Replace$(value, vbTab, "\t")
    value = Replace$(value, vbCr, "\r")
    value = Replace$(value, vbLf, "\n")
    Escape = value
End Function

'* @brief   Escape の逆変換
'* @param   value  エスケープ済み文字列
'* @return  元の文字列
Private Function Unescape(ByVal value As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String

    ' 大半のセルはエスケープを含まない。走査を省いて復元コストを抑える
    If InStr(value, "\") = 0 Then
        Unescape = value
        Exit Function
    End If

    i = 1
    Do While i <= Len(value)
        ch = Mid$(value, i, 1)
        If ch = "\" And i < Len(value) Then
            ' Replace の逆順では "\\t" が誤復元されるため1文字ずつ走査する
            Select Case Mid$(value, i + 1, 1)
            Case "\"
                result = result & "\"
            Case "t"
                result = result & vbTab
            Case "r"
                result = result & vbCr
            Case "n"
                result = result & vbLf
            Case Else
                result = result & Mid$(value, i, 2)
            End Select
            i = i + 2
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    Unescape = result
End Function
