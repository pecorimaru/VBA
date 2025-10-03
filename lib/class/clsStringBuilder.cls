Option Explicit

'******************************************************************************
' クラス名  : clsStringBuilder
' 機能      : 予め指定したメモリ領域をバッファとして利用する文字列連結クラス
' 説明      :
'   - VBA標準の文字列連結(&)は大規模テキストで非効率なため
'     StringBuilderパターンを採用
'
' バージョン: 1.0.0
' -----------------------------------------------------------------------------
' 履歴      : 2025/09/27  pecorimaru  Ver.1.0.0  新規作成
'******************************************************************************

'******************************************************************************
' バッファサイズ設定
' ※1文字あたり2バイトであるため、実際に占有するメモリ領域はバッファサイズ*2
'******************************************************************************
Private Const DEFAULT_BUFFER_SIZE As Long = 32768
Private Const GROWTH_FACTOR As Double = 2#         ' バッファ拡張倍率

'******************************************************************************
' モジュール名
'******************************************************************************
Private Const CLASS_NAME = "clsStringBuilder"

'******************************************************************************
' エラー番号
'******************************************************************************
Private Const ERR_NO_INDEX_NEGATIVE = vbObjectError + 1001
Private Const ERR_NO_INDEX_OUT_OF_RANGE = vbObjectError + 1002
Private Const ERR_NO_DELETE_LENGTH_NEGATIVE = vbObjectError + 1003

'******************************************************************************
' エラーメッセージ
'******************************************************************************
Private Const ERR_MSG_INDEX_NEGATIVE = "{0}位置が負の値です"
Private Const ERR_MSG_INDEX_OUT_OF_RANGE = "{0}位置（{1}）は範囲外です（可能範囲:0-{2}）"
Private Const ERR_MSG_DELETE_LENGTH_NEGATIVE = "削除文字数が負の値です"

'******************************************************************************
' クラス変数
'******************************************************************************
Private p_buf As String
Private p_currentLength As Long
Private p_bufferSizeOverCount As Long

'******************************************************************************
' 機　能    : クラスの初期化処理
'******************************************************************************
Private Sub Class_Initialize()
    p_buf = String$(DEFAULT_BUFFER_SIZE, vbNullChar)
    p_currentLength = 0
    p_bufferSizeOverCount = 0
End Sub

'******************************************************************************
' 機　能    : クラスの終了処理
'******************************************************************************
Private Sub Class_Terminate()
    If p_bufferSizeOverCount > 0 Then
        Debug.Print "バッファサイズ超過回数：" + CStr(p_bufferSizeOverCount) + "回"
        Debug.Print "最終文字数/バッファサイズ：" + CStr(Me.Length) + "/" + CStr(Me.Capacity)
    End If
    Call Dispose
End Sub

'******************************************************************************
' 機　能    : 現在の文字数を取得
'******************************************************************************
Public Property Get Length() As Long
    Length = p_currentLength
End Property

'******************************************************************************
' 機　能    : 現在のバッファサイズを取得
'******************************************************************************
Public Property Get Capacity() As Long
    Capacity = Len(p_buf)
End Property

'******************************************************************************
' 機　能    : バッファが空かどうかを判定
'******************************************************************************
Public Property Get IsEmpty() As Boolean
    IsEmpty = (p_currentLength = 0)
End Property

'******************************************************************************
' 機　能    : 全体の文字列を取得（Textプロパティ）
'******************************************************************************
Public Property Get Text() As String
    Text = Me.ToString()
End Property

'******************************************************************************
' 機　能    : 文字列をバッファに追加
' -----------------------------------------------------------------------------
' 引　数    : str - 追加する文字列
' 戻り値    : なし
'******************************************************************************
Public Sub Append(ByVal str As String)
    If str = "" Then
        Exit Sub
    End If
    Call EnsureCapacity(Len(str))
    Mid(p_buf, p_currentLength + 1) = str
    p_currentLength = p_currentLength + Len(str)
End Sub

'******************************************************************************
' 機　能    : バッファ内の全文字列を取得
' -----------------------------------------------------------------------------
' 引　数    : なし
' 戻り値    : String - バッファ内の全文字列（空の場合は空文字列）
'******************************************************************************
Public Function ToString() As String
    ToString = Left$(p_buf, p_currentLength)
End Function

'******************************************************************************
' 機　能    : 文字列と改行をバッファに追加
' 説  明    : 追加する文字列が省略された場合は改行のみを追加する
' -----------------------------------------------------------------------------
' 引　数    : str - 追加する文字列（省略可能）
' 戻り値    : なし
'******************************************************************************
Public Sub AppendLine(Optional ByVal str As String = "")
    Call Me.Append(str)
    Call Me.Append(vbCrLf)
End Sub

'******************************************************************************
' 機　能    : バッファの内容をクリア
' -----------------------------------------------------------------------------
' 引　数    : なし
' 戻り値    : なし
'******************************************************************************
Public Sub Clear()
    p_buf = String$(DEFAULT_BUFFER_SIZE, vbNullChar)
    p_currentLength = 0
End Sub

'******************************************************************************
' 機　能    : 指定位置に文字列を挿入
' -----------------------------------------------------------------------------
' 引　数    : offset - 挿入位置（0ベース）
'           : str    - 挿入する文字列
' 戻り値    : なし
'******************************************************************************
Public Sub Insert(ByVal offset As Long, ByVal str As String)
    ' 挿入位置の負数チェック
    If offset < 0 Then
        Err.Raise ERR_NO_INDEX_NEGATIVE, CLASS_NAME, FormatMessage(ERR_MSG_INDEX_NEGATIVE, "挿入")
    End If
    
    ' 挿入位置の範囲チェック
    If offset > p_currentLength Then
        Err.Raise ERR_NO_INDEX_OUT_OF_RANGE, CLASS_NAME, _
                FormatMessage(ERR_MSG_INDEX_OUT_OF_RANGE, "挿入", offset, p_currentLength)
    End If
    
    ' 挿入位置が末尾の場合はAppendを使用
    If offset = p_currentLength Then
        Call Me.Append(str)
        Exit Sub
    End If
    
    If str = "" Then
        Exit Sub
    End If

    Call EnsureCapacity(Len(str))
    
    ' 挿入位置以降の文字列を挿入する文字数分シフトする
    Mid(p_buf, offset + 1 + Len(str)) = Mid$(p_buf, offset + 1, p_currentLength)
    
    ' 挿入位置に文字列をセット
    Mid(p_buf, offset + 1) = str
    
    p_currentLength = p_currentLength + Len(str)

End Sub

'******************************************************************************
' 機　能    : 指定範囲の文字列を削除
' -----------------------------------------------------------------------------
' 引　数    : start        - 削除開始位置（0ベース）
'           : removeLength - 削除する文字数
' 戻り値    : なし
'******************************************************************************
Public Sub Remove(ByVal start As Long, ByVal removeLength As Long)
    ' 削除開始位置の負数チェック
    If start < 0 Then
        Err.Raise ERR_NO_INDEX_NEGATIVE, CLASS_NAME, FormatMessage(ERR_MSG_INDEX_NEGATIVE, "削除開始")
    End If
    
    ' 削除開始位置の範囲チェック
    If start >= p_currentLength Then
        Err.Raise ERR_NO_INDEX_OUT_OF_RANGE, CLASS_NAME, _
                FormatMessage(ERR_MSG_INDEX_OUT_OF_RANGE, "削除開始", start, p_currentLength - 1)
    End If
    
    ' 削除文字数の負数チェック
    If removeLength < 0 Then
        Err.Raise ERR_NO_DELETE_LENGTH_NEGATIVE, CLASS_NAME, ERR_MSG_DELETE_LENGTH_NEGATIVE
    End If
    
    If removeLength = 0 Then
        Exit Sub
    End If

    ' 削除範囲を調整（バッファサイズを超えないように）
    Dim actualLength As Long
    If start + removeLength > p_currentLength Then
        actualLength = p_currentLength - start
    Else
        actualLength = removeLength
    End If
    
    ' 削除開始位置 + 削除する文字数以降の文字列を削除開始位置へシフトする
    Mid(p_buf, start + 1) = Mid$(p_buf, start + actualLength + 1, p_currentLength - actualLength)
    
    ' バッファ内の全文字数 - 削除する文字数以降の文字列をクリアする
    Mid(p_buf, p_currentLength - actualLength + 1) = String$(removeLength, vbNullChar)
    
    p_currentLength = p_currentLength - actualLength

End Sub

'******************************************************************************
' 機　能    : バッファサイズの拡張
' 説　明    : バッファサイズを超過した場合はサイズを拡張して再配置する
'           : -- サイズの拡張方法 --
'           : 現在のバッファサイズ + (初期バッファサイズ * 超過回数 * 拡張倍率) + 追加文字数
' -----------------------------------------------------------------------------------
' 引　数    : strLength - 追加文字数
' 戻り値    : なし
'******************************************************************************
Private Sub EnsureCapacity(ByVal strLength As Long)
    If p_currentLength + strLength > Len(p_buf) Then
        p_bufferSizeOverCount = p_bufferSizeOverCount + 1
        p_buf = p_buf & String$(CLng(DEFAULT_BUFFER_SIZE * p_bufferSizeOverCount * GROWTH_FACTOR) + strLength, vbNullChar)
    End If
End Sub

'******************************************************************************
' 機　能    : リソースの解放処理
' -----------------------------------------------------------------------------
' 引　数    : なし
' 戻り値    : なし
'******************************************************************************
Private Sub Dispose()
    p_buf = vbNullString
    p_currentLength = 0
    p_bufferSizeOverCount = 0
End Sub

'******************************************************************************
' 機　能    : プレースホルダーを置き換えた文字列を取得
' 説  明    : テンプレート文字列内の{0},{1}...を指定された値で置き換える
' 注意事項  : ・プレースホルダーは{0}から開始（0ベース）
'           : ・引数が不足した場合、該当プレースホルダーはそのまま残る
' -----------------------------------------------------------------------------
' 引　数    : template - テンプレート文字列
'           : args     - 置き換える値（可変引数）
' 戻り値    : String   - 置き換え後の文字列
'******************************************************************************
Private Function FormatMessage( _
    ByVal template As String, _
    ParamArray args() As Variant _
) As String
    
    Dim result As String
    Dim i As Integer
    
    result = template
    
    ' 各引数でプレースホルダーを置換
    For i = 0 To UBound(args)
        result = Replace(result, "{" & CStr(i) & "}", CStr(args(i)))
    Next i
    
    FormatMessage = result
    
End Function

