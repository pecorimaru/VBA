Attribute VB_Name = "modAppProc"
Option Explicit

'************************************************************************************
' 機　能    :アプリケーション操作モジュール
' 依存関係  :  shMsg
'           :  modExclEdit
'           :  modFileUtils
'           :  modStringUtils
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/11/01  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************

'----------グローバル定数----------

'メッセージボックス用
Public Enum XlMsgType
    xlInfo      'インフォメーション、通知
    xlWarning   'ワーニング、警告
    xlError     'エラー
End Enum

'----------messageシート----------

'行Index
Private Enum ROW_SH_MSG
    IDX_HDR = 1
    IDX_DTL_BGN = 2
End Enum

'メッセージ 列Index
Private Enum COL_SH_MSG
    IDX_ID = 1
    IDX_MSG = 2
End Enum

'----------モジュール変数----------

'アプリケーション設定 保存用
Private Type AppSettingsType
    enableEventsValue As Boolean
    interactiveValue As Boolean
    screenUpdatingValue As Boolean
    cursorValue As XlMousePointer
    calculationValue As XlCalculation
    displayStatusBarValue As Boolean
    statusBarValue As Variant
End Type

'アプリケーション設定 保存用
Private appSettingsSave As AppSettingsType         '処理開始前
Private appSettingsPauseSave As AppSettingsType    '一時停止前

'アプリケーション設定 フラグ
Private isProcessing As Boolean
Private isPausing As Boolean

'************************************************************************************
' 機　能    :メッセージボックス表示
' 引　数    :in      prompt                     メッセージ内容
'           :in      msgType                    メッセージ種別
'           :                                     ・xlInfo   (インフォメーション、通知)
'           :                                     ・xlWarning(ワーニング、警告)
'           :                                     ・xlError  (エラー)
'           :in      buttons                    ボタンやアイコン種類
'           :in      ttl                        タイトル
'           :in      withAutoCalc               再計算フラグ
' 戻　値    :押下されたボタンの種類
'************************************************************************************
Public Function showMsgBox( _
      ByVal prompt As String _
    , ByRef msgType As XlMsgType _
    , Optional ByVal buttons As VbMsgBoxStyle = VbMsgBoxStyle.vbOKOnly _
    , Optional ByVal ttl As String = vbNullString _
    , Optional ByVal withAutoCalc As Boolean = True _
) As VbMsgBoxResult

    On Error GoTo ErrorHandler

    '処理を一時中断
    Call setApplicationSettingsPause(False)
    
    '再計算
    If withAutoCalc Then
        Application.Calculate
    End If

    'buttonsパラメータからアイコン指定を抽出
    Dim wkButtons As Long
    Dim icon As Long
    wkButtons = buttons
    icon = (wkButtons And &HF0)
    
    'アイコン指定ありの場合
    If icon > 0 Then
        'アイコン指定を削除
        wkButtons = wkButtons - icon
    End If
    
    'メッセージ種別に合わせてアイコンを指定
    Select Case msgType
        Case XlMsgType.xlInfo
            wkButtons = wkButtons + VbMsgBoxStyle.vbInformation
        Case XlMsgType.xlWarning
            wkButtons = wkButtons + VbMsgBoxStyle.vbExclamation
        Case XlMsgType.xlError
            wkButtons = wkButtons + VbMsgBoxStyle.vbCritical
    End Select
    
    'タイトルが指定されなかった場合
    If StrPtr(ttl) = 0 Then
        'ワークブック名をタイトルに指定
        ttl = getFileNameNoneExtFromPath(ThisWorkbook.Name)
    End If

    showMsgBox = MsgBox(prompt, wkButtons, ttl)

    GoTo Fin

ErrorHandler:
    'エラー処理

Fin:
    '処理を再開
    Call setApplicationSettingsReStart

    'エラーが発生した場合は呼出元に投げる
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If

End Function

'************************************************************************************
' 機　能    :システムエラーメッセージを表示する
' 引　数    :in      errObj                     エラーオブジェクト
'           :in      ttl                        タイトル
' 戻　値    :なし
'************************************************************************************
Public Sub showSystemErrorMsg( _
      ByRef errObj As ErrObject _
    , Optional ByVal ttl As String = vbNullString _
)
    Call showMsgBox( _
              "C_ME0001" & vbCrLf & errObj.Description & "(" & CStr(errObj.Number) & ")" _
            , XlMsgType.xlError _
            , VbMsgBoxStyle.vbOKOnly _
            , ttl _
         )
End Sub

'************************************************************************************
' 機　能    :メッセージ取得
' 引　数    :in  id                             メッセージＩＤ
'           :in  replaceStrings                 置き換え文字列(複数指定可)
' 戻　値    :メッセージ
'************************************************************************************
Public Function getMsg( _
      ByVal id As String _
    , ParamArray replaceStrings()) As String

    On Error GoTo ErrorHandler

    'メッセージシートの最大行取得
    Dim rowMax As Long
    rowMax = getRowMax(shMsg, COL_SH_MSG.IDX_ID)
    If rowMax < ROW_SH_MSG.IDX_DTL_BGN Then
        getMsg = ""
        GoTo Fin
    End If

    'メッセージシートの内容を二次元配列に取得
    Dim msgAry As Variant
    msgAry = cnvRange2Ary(shMsg, ROW_SH_MSG.IDX_DTL_BGN, rowMax, COL_SH_MSG.IDX_ID, COL_SH_MSG.IDX_MSG)
        
    'メッセージクリア
    Dim msg As String
    msg = ""

    '配列をサーチしてＩＤが一致する場合にメッセージを取得
    Dim row As Long
    For row = LBound(msgAry) To UBound(msgAry)
        If Trim(msgAry(row, COL_SH_MSG.IDX_ID)) = id Then
            msg = msgAry(row, COL_SH_MSG.IDX_MSG)
            Exit For
        End If
    Next

    'メッセージが取得できた場合、置き換え文字列を置換
    If Not isBlank(msg) Then
        msg = getReplacedMessage(msg, replaceStrings)
    End If

    'メッセージを返す
    getMsg = msg & "(" & id & ")"

    GoTo Fin

ErrorHandler:
    'エラー処理

Fin:
    '後処理
    If Not IsEmpty(msgAry) Then
        Erase msgAry
    End If

    'エラーが発生した場合は呼出元に投げる
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If

End Function

'************************************************************************************
' 機　能    :メッセージの置き換え文字列を置換
' 引　数    :in  msg                            メッセージ
'           :in  replaceStrings                 置き換え文字列配列
' 戻　値    :メッセージ
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/06/14  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************
Private Function getReplacedMessage( _
      ByVal msg As String _
    , ByVal replaceStrings As Variant _
) As String

    Dim wkMsg As String
    wkMsg = msg
    
    Dim i As Integer
    i = 0
    
    Dim rep As Variant
    For Each rep In replaceStrings
        i = i + 1
        wkMsg = Replace(wkMsg, "{" & CStr(i) & "}", CStr(rep))
    Next

    getReplacedMessage = wkMsg

End Function

'************************************************************************************
' 機　能    :エラーオブジェクトRaise
' 引　数    :in  msg                            メッセージ
'           :in  errSource                      エラー発生箇所
'           :in  replaceStrings                 メッセージ置き換え文字列(複数指定可)
' 戻　値    :エラーオブジェクト
'************************************************************************************
Public Sub raiseCommonErr( _
      ByVal msg As String _
    , ByVal errSource As String _
    , ParamArray replaceStrings() _
)

    '＜前提＞
    'メッセージ …「メッセージID:メッセージ内容」形式
    'メッセージＩＤ …「C_Mx0000」形式

    'メッセージより、Err.Numberを取得する(メッセージIDの数値部分 + vbObjectError)
    Dim msgNo As Long
    msgNo = CLng(Mid(msg, 5, 4))
    msgNo = msgNo + vbObjectError

    'エラーRaise
    Call Err.Raise(msgNo, errSource, ReplacePh(msg, replaceStrings))

End Sub

'************************************************************************************
' 機　能    :アプリケーション設定(処理開始)
' 説　明    :VBA実行中に画面描画やユーザー操作が行われないように制御
' 引　数    :in      statusbarString            ステータスバー表示文字列
'           :                                     ※"FALSE"という文字列は指定不可
' 戻　値    :なし
'************************************************************************************
Public Sub setApplicationSettingsStart( _
    Optional ByVal statusbarString As String = vbNullString _
)

    If Not isProcessing Then
    
        '処理開始前の設定を保存
        appSettingsSave = getAppSettings

        With Application
            
            '設定
            .EnableEvents = False                               'イベントを抑制
            .Interactive = False                                'ユーザ操作を受け付けない
            .ScreenUpdating = False                             '描画更新を停止
            .Calculation = XlCalculation.xlCalculationManual    '数式再計算を手動に設定
            .Cursor = XlMousePointer.xlWait                     'カーソルをWAITに変更

            'ステータスバー表示文字列が指定された場合
            If Not StrPtr(statusbarString) = 0 Then
                'ステータスバーを変更
                .DisplayStatusBar = True
                .StatusBar = statusbarString
            End If
        
        End With

        isProcessing = True
        isPausing = False
        
    End If

End Sub

'************************************************************************************
' 機　能    :アプリケーション設定(処理終了)
' 説　明    :設定を処理開始前に戻す
' 引　数    :なし
' 戻　値    :なし
'************************************************************************************
Public Sub setApplicationSettingsEnd()

    If isProcessing Then
    
        '設定を処理開始前に戻す
        '  ステータスバー：変更する(処理開始前の状態に戻す)
        '  計算方法：変更する(処理開始前の状態に戻す)
        Call setAppSettings(appSettingsSave, True, True)

        isProcessing = False
        isPausing = False
    
    End If

End Sub

'************************************************************************************
' 機　能    :アプリケーション設定(一時停止)
' 引　数    :in      calcModeChange             計算方法を処理開始前の状態に変更する
' 戻　値    :なし
'************************************************************************************
Public Sub setApplicationSettingsPause( _
    Optional ByVal calcModeChange As Boolean = True _
)

    If (isProcessing) And (Not isPausing) Then
    
        '一時停止前の設定を保存
        appSettingsPauseSave = getAppSettings

        '設定を処理開始前に戻す
        '  ステータスバー：変更しない
        '  計算方法：処理開始前の状態に変更するかどうかフラグで判断
        Call setAppSettings(appSettingsSave, False, calcModeChange)

        '一時停止フラグＯＮ
        isPausing = True
    
    End If

End Sub

'************************************************************************************
' 機　能    :アプリケーション設定(再始動)
' 引　数    :なし
' 戻　値    :なし
'************************************************************************************
Public Sub setApplicationSettingsReStart()

    If isPausing Then

         '設定を一時停止前に戻す
         '  ステータスバー：変更しない
         '  計算方法：変更する(一時停止前の状態に戻す)
         Call setAppSettings(appSettingsPauseSave, False, True)
 
         '一時停止フラグＯＦＦ
         isPausing = False

    End If

End Sub

'************************************************************************************
' 機　能    :アプリケーション設定(画面更新)
' 引　数    :in      withAutoCalc               再計算フラグ
' 戻　値    :なし
'************************************************************************************
Public Sub setApplicationSettingsRefresh( _
    Optional ByVal withAutoCalc As Boolean = True _
)
        
    '現在の設定を保存
    Dim tmpAppSettingsSave As AppSettingsType
    tmpAppSettingsSave = getAppSettings

    With Application
        
        '画面更新
        .EnableEvents = True
        .Interactive = True
        .ScreenUpdating = True

        '再計算する
        If withAutoCalc Then
            .Calculate
        End If
    
    End With

    DoEvents

    '設定を戻す
    '  ステータスバー：変更しない
    '  計算方法：変更しない
    Call setAppSettings(tmpAppSettingsSave, False, False)

End Sub

'************************************************************************************
' 機　能    :現在のアプリケーション設定を取得
' 引　数    :なし
' 戻　値    :現在のアプリケーション設定
'************************************************************************************
Private Function getAppSettings() As AppSettingsType

    Dim rtn As AppSettingsType
    
    With Application
        rtn.enableEventsValue = .EnableEvents
        rtn.interactiveValue = .Interactive
        rtn.screenUpdatingValue = .ScreenUpdating
        rtn.cursorValue = .Cursor
        rtn.calculationValue = .Calculation
        rtn.displayStatusBarValue = .DisplayStatusBar
        rtn.statusBarValue = .StatusBar
    End With

    getAppSettings = rtn

End Function

'************************************************************************************
' 機　能    :アプリケーション設定を変更
' 引　数    :in      appSettingsValue           アプリケーション設定値
'           :in      statusBarSet               ステータスバーを設定するか
'           :in      calcModeSet                計算方法を設定するか
' 戻　値    :なし
'************************************************************************************
Private Sub setAppSettings( _
      ByRef appSettingsValue As AppSettingsType _
    , ByVal statusBarSet As Boolean _
    , ByVal calcModeSet As Boolean _
)

    With Application
        .EnableEvents = appSettingsValue.enableEventsValue
        .Interactive = appSettingsValue.interactiveValue
        .ScreenUpdating = appSettingsValue.screenUpdatingValue
        .Cursor = appSettingsValue.cursorValue

        If calcModeSet Then
            .Calculation = appSettingsValue.calculationValue
        End If

        If statusBarSet Then
            If UCase(appSettingsValue.statusBarValue) = "FALSE" Then
                .StatusBar = False
                .DisplayStatusBar = appSettingsValue.displayStatusBarValue
            Else
                .StatusBar = appSettingsValue.statusBarValue
                .DisplayStatusBar = appSettingsValue.displayStatusBarValue
            End If
        End If
    End With

End Sub



