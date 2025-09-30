VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStringBuilderFileBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'******************************************************************************
' クラス名  : StringBuilder（ファイルベース実装）
' 機能      : 一時ファイルをバッファとして利用する文字列連結クラス
' 説明      :
'   - VBA標準の文字列連結(&)は大規模テキストで非効率なため、
'     StringBuilderパターンを採用。
'   - メモリ不足対策として、一時ファイルをバッファに使用。
'   - 大容量テキスト処理やログ生成に適しています。
'
' バージョン: 1.0.0
' -----------------------------------------------------------------------------
' 履歴      : 2025/09/27  Pecorimaru  Ver.1.0.0  新規作成
'******************************************************************************

'******************************************************************************
' モジュール名
'******************************************************************************
Private Const CLASS_NAME = "clsStringBuilderFileBuffer"

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
Private Const ERR_MSG_TEMP_FILE_CREATE_FAILED = "一時ファイルの作成に失敗しました（{0}回リトライ後）: {1}"

'******************************************************************************
' クラス変数
'******************************************************************************
Private p_tempFilePath As String        ' 一時ファイルのパス
Private p_fileHandle As Integer         ' ファイルハンドル
Private p_currentLength As Long         ' 現在の文字数
Private p_isFileOpen As Boolean         ' ファイルオープン状態

'******************************************************************************
' 機　能    : クラスの初期化処理
'******************************************************************************
Private Sub Class_Initialize()
    Call CreateTempFile
    p_currentLength = 0
End Sub

'******************************************************************************
' 機　能    : クラスの終了処理
'******************************************************************************
Private Sub Class_Terminate()
    Call Dispose
End Sub

'******************************************************************************
' 機　能    : 現在の文字数を取得
'******************************************************************************
Public Property Get Length() As Long
    Length = p_currentLength
End Property

'******************************************************************************
' 機　能    : バッファが空かどうかを判定
'******************************************************************************
Public Property Get IsEmpty() As Boolean
    IsEmpty = (p_currentLength = 0)
End Property

'******************************************************************************
' 機　能    : 全体の文字列を取得
'******************************************************************************
Public Property Get Text() As String
    Text = Me.ToString()
End Property

'******************************************************************************
' 機　能    : 文字列をバッファに追加
' 説  明    : 指定された文字列を一時ファイルの末尾に追加する
' -----------------------------------------------------------------------------
' 引　数    : text - 追加する文字列
' 戻り値    : なし
' 例　外    : ファイルが開かれていない場合はエラーが発生
'******************************************************************************
Public Sub Append(ByVal addText As String)
    If addText = "" Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Print #p_fileHandle, addText;
    p_currentLength = p_currentLength + Len(addText)
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, CLASS_NAME & ".Append", Err.Description
End Sub

'******************************************************************************
' 機　能    : 文字列と改行をバッファに追加
' 説  明    : 指定された文字列の末尾に改行文字（vbCrLf）を付加してバッファに追加する
'           : textが省略された場合は改行のみを追加する
' -----------------------------------------------------------------------------
' 引　数    : text - 追加する文字列（省略可能、省略時は空文字列）
' 戻り値    : なし
' 例　外    : ファイルが開かれていない場合はエラーが発生
'******************************************************************************
Public Sub AppendLine(Optional ByVal addText As String = "")
    Me.Append addText
    Me.Append vbCrLf
End Sub

'******************************************************************************
' 機　能    : バッファ内の全文字列を取得
' 説  明    : 一時ファイルに蓄積された全ての文字列を読み込んで返す
' 注意事項  : ・ファイル開閉処理を伴うため、頻繁な呼び出しは避ける
'           : ・大容量データの場合はメモリ不足の可能性あり
' -----------------------------------------------------------------------------
' 引　数    : なし
' 戻り値    : String - バッファ内の全文字列（空の場合は空文字列）
' 例外      : ファイル読み取りエラー時は例外が発生
'******************************************************************************
Public Function ToString() As String
    If p_currentLength = 0 Then
        ToString = ""
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    Dim readHandle As Integer
    Dim fileContent As String
    Dim originalHandle As Integer
    
    ' 現在のハンドルを保存
    originalHandle = p_fileHandle
    
    ' 現在のファイルをフラッシュして閉じる
    Close #p_fileHandle
    p_isFileOpen = False
    
    ' 読み取り専用で開く
    readHandle = FreeFile
    Open p_tempFilePath For Binary As #readHandle
    
    ' ファイル全体を読み込み
    If LOF(readHandle) > 0 Then
        fileContent = Input$(LOF(readHandle), #readHandle)
    Else
        fileContent = ""
    End If
    
    Close #readHandle
    
    ' 元のファイルを書き込みモードで再度開く
    p_fileHandle = FreeFile
    Open p_tempFilePath For Append As #p_fileHandle
    p_isFileOpen = True
    
    ToString = fileContent
    Exit Function
    
ErrorHandler:
    ' ファイルが開かれている場合は強制的に閉じる
    On Error Resume Next
    If readHandle > 0 Then Close #readHandle
    Err.Raise Err.Number, CLASS_NAME & ".ToString", Err.Description
End Function

'************************************************************************************
' 機　能    : バッファの内容をクリア
' 説  明    : 一時ファイルを再作成してバッファ内の全ての文字列を削除する
' -----------------------------------------------------------------------------------
' 引　数    : なし
' 戻り値    : なし
' 例外      : ファイル操作エラー時は例外が発生
'************************************************************************************
Public Sub Clear()
    
    On Error GoTo ErrorHandler
    
    Close #p_fileHandle
    Open p_tempFilePath For Output As #p_fileHandle
    p_currentLength = 0
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, CLASS_NAME & ".Clear", Err.Description
End Sub

'******************************************************************************
' 機　能    : 指定位置に文字列を挿入
' 説  明    : バッファ内の指定された位置に文字列を挿入する
'           : ファイルベース実装のため、一時的にメモリに読み込んで再構築する
' 注意事項  : ・大容量データの場合はメモリ不足の可能性あり
'           : ・パフォーマンスが悪いため、頻繁な使用は避ける
'           : ・挿入位置が範囲外の場合はエラーが発生
' -----------------------------------------------------------------------------
' 引　数    : index - 挿入位置（0ベース）
'           : text  - 挿入する文字列
' 戻り値    : なし
' 例　外    : ファイル操作エラー、インデックス範囲外エラーが発生
'******************************************************************************
Public Sub Insert(ByVal index As Long, ByVal addText As String)
    
    ' 挿入位置の負数チェック
    If index < 0 Then
        Err.Raise ERR_NO_INDEX_NEGATIVE, CLASS_NAME, FormatMessage(ERR_MSG_INDEX_NEGATIVE, "挿入")
    End If
    
    ' 挿入位置の範囲チェック
    If index > p_currentLength Then
        Err.Raise ERR_NO_INDEX_OUT_OF_RANGE, CLASS_NAME, _
                FormatMessage(ERR_MSG_INDEX_OUT_OF_RANGE, "挿入", index, p_currentLength)
    End If
    
    ' 空文字列の場合は何もしない
    If addText = "" Then
        Exit Sub
    End If
    
    ' 末尾への挿入の場合はAppendを使用（最適化）
    If index = p_currentLength Then
        Call Me.Append(addText)
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Dim originalContent As String
    Dim beforePart As String
    Dim afterPart As String
    Dim newContent As String
    
    ' 現在の内容を取得
    originalContent = Me.ToString()
    
    ' 挿入位置で文字列を分割
    beforePart = Left(originalContent, index)
    afterPart = Mid(originalContent, index + 1)
        
    ' ファイルをクリアして新しい内容を書き込み
    Me.Clear
    Me.Append beforePart
    Me.Append addText
    Me.Append afterPart
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, CLASS_NAME & ".Insert", Err.Description
End Sub

'******************************************************************************
' 機　能    : 指定範囲の文字列を削除
' 説  明    : バッファ内の指定された範囲の文字列を削除する
'           : ファイルベース実装のため、一時的にメモリに読み込んで再構築する
' 注意事項  : ・大容量データの場合はメモリ不足の可能性あり
'           : ・パフォーマンスが悪いため、頻繁な使用は避ける
'           : ・削除範囲が範囲外の場合はエラーが発生
' -----------------------------------------------------------------------------
' 引　数    : startIndex - 削除開始位置（0ベース）
'           : length     - 削除する文字数
' 戻り値    : なし
' 例　外    : ファイル操作エラー、インデックス範囲外エラーが発生
'******************************************************************************
Public Sub Remove(ByVal startIndex As Long, ByVal removeLength As Long)
    ' 削除開始位置の負数チェック
    If startIndex < 0 Then
        Err.Raise ERR_NO_INDEX_NEGATIVE, CLASS_NAME, FormatMessage(ERR_MSG_INDEX_NEGATIVE, "削除開始")
    End If
    
    ' 削除開始位置の範囲チェック
    If startIndex >= p_currentLength Then
        Err.Raise ERR_NO_INDEX_OUT_OF_RANGE, CLASS_NAME, _
                FormatMessage(ERR_MSG_INDEX_OUT_OF_RANGE, "削除開始", startIndex, p_currentLength - 1)
    End If
    
    ' 削除文字数の負数チェック
    If removeLength < 0 Then
        Err.Raise ERR_NO_DELETE_LENGTH_NEGATIVE, CLASS_NAME, ERR_MSG_DELETE_LENGTH_NEGATIVE
    End If
        
    ' 削除文字数が0の場合は何もしない
    If removeLength = 0 Then
        Exit Sub
    End If
    
    ' 削除範囲を調整（バッファサイズを超えないように）
    Dim actualLength As Long
    If startIndex + removeLength > p_currentLength Then
        actualLength = p_currentLength - startIndex
    Else
        actualLength = removeLength
    End If
    
    On Error GoTo ErrorHandler
    
    ' 現在の内容を取得
    Dim originalContent As String
    originalContent = Me.ToString()
    
    ' 削除範囲の終了位置を計算
    Dim endIndex As Long
    endIndex = startIndex + actualLength
    
    ' 削除前後の部分を取得
    Dim beforePart As String
    beforePart = Left(originalContent, startIndex)
    
    Dim afterPart As String
    afterPart = Mid(originalContent, endIndex + 1)
        
    ' ファイルをクリアして新しい内容を書き込み
    Me.Clear
    Me.Append beforePart
    Me.Append afterPart
        
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, CLASS_NAME & ".Remove", Err.Description
End Sub

'************************************************************************************
' 機　能    : 一時ファイルを作成してバッファを初期化
' 説  明    : TEMPフォルダに一意な名前の一時ファイルを作成し、書き込み用に開く
' 注意事項  : ・ファイル名は日時とランダム値で重複を回避
'           : ・作成失敗時は最大3回までリトライする
'           : ・全てのリトライが失敗した場合はp_isFileOpenがFalseに設定される
' -----------------------------------------------------------------------------------
' 引　数    : なし
' 戻り値    : なし
' 例外      : 全てのリトライが失敗した場合は例外が発生
'************************************************************************************
Private Sub CreateTempFile()
    Const MAX_RETRY_COUNT As Integer = 3
    Dim retryCount As Integer
    Dim lastError As String
    
    On Error GoTo ErrorHandler
    
    For retryCount = 1 To MAX_RETRY_COUNT
        ' 一時ファイルパスを生成（リトライ毎に異なる名前）
        p_tempFilePath = Environ("TEMP") & "\" & "SB_" & Format(Now, "yyyymmdd_hhnnss") & "_" & Int(Rnd * 10000) & "_" & retryCount & ".tmp"
        
        ' ファイルハンドルを取得
        p_fileHandle = FreeFile
        
        ' ファイルを開く（出力モードで作成）
        Open p_tempFilePath For Output As #p_fileHandle
        p_isFileOpen = True
        
        Exit Sub
        
ErrorHandler:
        lastError = Err.Description
        
        ' ファイルが開いていたら閉じる
        On Error Resume Next
        If p_fileHandle > 0 Then
            Close #p_fileHandle
        End If
        On Error GoTo ErrorHandler
        
        ' 最後のリトライでも失敗した場合
        If retryCount = MAX_RETRY_COUNT Then
            p_isFileOpen = False
            Err.Raise Err.Number, CLASS_NAME & ".CreateTempFile", _
                    Replace(ERR_MSG_TEMP_FILE_CREATE_FAILED, MAX_RETRY_COUNT, lastError)
        End If
        
        ' 少し待ってからリトライ（オプション）
        Application.Wait DateAdd("s", 1, Now)
    Next retryCount
End Sub

'************************************************************************************
' 機　能    : リソースの解放処理
' 説  明    : ファイルハンドルの解放、一時ファイルの削除、内部変数の初期化を行う
' 注意事項  : ・全ての処理でエラーを無視する（On Error Resume Next）
'           : ・Class_Terminateから自動的に呼び出される
' -----------------------------------------------------------------------------------
' 引　数    : なし
' 戻り値    : なし
' 例外      : なし（全てのエラーを無視）
'************************************************************************************
Public Sub Dispose()
    On Error Resume Next
    
    ' ファイルを閉じる
    If p_isFileOpen And p_fileHandle > 0 Then
        Close #p_fileHandle
        p_isFileOpen = False
    End If
    
    ' 一時ファイルを削除
    If p_tempFilePath <> "" Then
        If Dir(p_tempFilePath) <> "" Then
            Kill p_tempFilePath
        End If
        p_tempFilePath = ""
    End If
    
    ' 変数をリセット
    p_fileHandle = 0
    p_currentLength = 0
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
' 例　外    : なし
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
 
