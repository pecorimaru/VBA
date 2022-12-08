Attribute VB_Name = "modFileUtils"
Option Explicit

'************************************************************************************
' 機　能    :ファイル操作の汎用モジュール
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/11/01  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************

'************************************************************************************
' 機　能    :ファイル選択ダイアログ（汎用：フィルタ１種類／１件選択）
' 引　数    :in  ttl                             ダイアログのタイトル
'           :in  filterName                      フィルタ名
'           :in  filterExtension                 フィルタ拡張子
'           :in  initialFileName                 初期選択値
' 戻　値    :選択ファイル名
'************************************************************************************
Public Function showFilePickerDialog( _
      ByVal ttl As String _
    , ByVal filterName As String _
    , ByVal filterExtension As String _
    , Optional ByVal initialFileName As String = "" _
) As String

    '戻り値をクリア
    showFilePickerDialog = ""

    'ダイアログ情報
    With Application.FileDialog(msoFileDialogOpen)

        'ダイアログのタイトル
        .title = ttl

        'ダイアログのフィルタ
        .Filters.Clear
        .Filters.add filterName, filterExtension
        .FilterIndex = 1

        '複数ファイル選択を許可しない
        .AllowMultiSelect = False

        'ダイアログの初期選択値
        If Trim(initialFileName) <> "" Then
            .initialFileName = initialFileName
        End If

        'ダイアログを表示
        If .Show <> 0 Then
            If .SelectedItems.count > 0 Then
                'ファイルが選択された場合：戻り値にセット
                showFilePickerDialog = .SelectedItems.item(1)
            End If
        End If
    
    End With

End Function

'************************************************************************************
' 機　能    :ファイル選択ダイアログ（汎用：フィルタ１種類／複数選択可能）
' 引　数    :in  ttl                             ダイアログのタイトル
'           :in  filterName                      フィルタ名
'           :in  filterExtension                 フィルタ拡張子
'           :in  initialFileName                 初期選択値
' 戻　値    :選択ファイルパスリスト
'************************************************************************************
Public Function showFilesPickerDialog( _
      ByVal ttl As String _
    , ByVal filterName As String _
    , ByVal filterExtension As String _
    , Optional ByVal initialFileName As String = "" _
) As String()

    Dim selItemList() As String

    'ダイアログ情報
    With Application.FileDialog(msoFileDialogOpen)

        'ダイアログのタイトル
        .title = ttl

        'ダイアログのフィルタ
        .Filters.Clear
        .Filters.add filterName, filterExtension
        .FilterIndex = 1

        '複数選択を許可
        .AllowMultiSelect = True

        'ダイアログの初期選択値
        If Trim(initialFileName) <> "" Then
            .initialFileName = initialFileName
        End If

        'ダイアログを表示
        If .Show <> 0 Then
            If .SelectedItems.count > 0 Then
                
                '選択ファイルリスト
                ReDim selItemList(0) As String
                
                '選択したファイル数分ループ
                Dim selItem As Variant
                For Each selItem In .SelectedItems
                                        
                    '配列の要素が存在する場合
                    If selItemList(0) <> "" Then
                     
                         '配列を拡張
                        ReDim Preserve selItemList(UBound(selItemList) + 1)
                     
                    End If
                                         
                    '選択ファイルを要素に追加
                    selItemList(UBound(selItemList)) = selItem
                     
                 Next selItem
            
            End If
        End If
    
    End With

    '選択ファイルパスリストを返却
    showFilesPickerDialog = selItemList

End Function

'************************************************************************************
' 機　能    :フォルダ選択ダイアログ（汎用）
' 引　数    :in  title                           ダイアログのタイトル
'           :in  initialFolderPath               フィルタ名
' 戻　値    :選択フォルダ名
'************************************************************************************
Public Function showFolderPickerDialog( _
      ByVal title As String _
    , Optional ByVal initialFolderPath As String = "" _
) As String
    
    '戻り値をクリア
    showFolderPickerDialog = ""

    'フォルダ選択ダイアログ設定
    With Application.FileDialog(msoFileDialogFolderPicker)
        
        'ダイアログのタイトル
        .title = title
        
        'ダイアログの初期選択値
        If Trim(initialFolderPath) <> "" Then
            .initialFileName = initialFolderPath
        End If

        'ダイアログを表示
        If .Show <> 0 Then
            If .SelectedItems.count > 0 Then
                'フォルダが選択された場合：戻り値にセット
                showFolderPickerDialog = .SelectedItems.item(1)
            End If
        End If
    
    End With

End Function

'************************************************************************************
' 機　能    :ファイル名を取得(拡張子を含む)
' 引　数    :in      path                       ファイルパス
' 戻　値    :ファイル名(拡張子を含む)
'************************************************************************************
Public Function getFileNameFromPath(ByVal path As String) As String

    Dim pos As Long

    pos = InStrRev(path, "\")

    If pos = 0 Then
        getFileNameFromPath = path
    Else
        getFileNameFromPath = Mid(path, pos + 1)
    End If

End Function

'************************************************************************************
' 機　能    :ファイル名を取得(拡張子を含まない)
' 引　数    :in      path                       ファイルパス
' 戻　値    :ファイル名(拡張子を含まない)
'************************************************************************************
Public Function getFileNameNoneExtFromPath(ByVal path As String) As String

    Dim fileName As String
    fileName = getFileNameFromPath(path)

    Dim pos As Long
    pos = InStrRev(fileName, ".")

    If pos = 0 Then
        getFileNameNoneExtFromPath = fileName
    Else
        getFileNameNoneExtFromPath = Mid(fileName, 1, pos - 1)
    End If

End Function

'************************************************************************************
' 機　能    :拡張子を取得
' 引　数    :in      path                       ファイルパス
' 戻　値    :拡張子
'************************************************************************************
Public Function getExtensionName(ByVal path As String) As String

    Dim fileName As String
    fileName = getFileNameFromPath(path)

    Dim pos As Long
    pos = InStrRev(fileName, ".")

    If pos = 0 Then
        getExtensionName = ""
    Else
        getExtensionName = Right(fileName, Len(fileName) - pos)
    End If

End Function

'************************************************************************************
' 機　能    :フォルダパスを取得
' 引　数    :in      path                       ファイルパス
' 戻　値    :フォルダ名パス
'************************************************************************************
Public Function getFolderFromPath(ByVal path As String) As String

    Dim wkPath As String
    wkPath = path

    '末尾が"\"の場合、削除
    If Right(wkPath, 1) = "\" Then
        wkPath = Mid(wkPath, 1, Len(wkPath) - 1)
    End If

    Dim pos As Long
    pos = InStrRev(wkPath, "\")

    If pos = 0 Then
        getFolderFromPath = path
    Else
        getFolderFromPath = Mid(wkPath, 1, pos - 1)
    End If

End Function

'************************************************************************************
' 機　能    :ファイル存在チェック
' 引　数    :in      path                       ファイルパス
'           :in      fso                        FileSystemObject
' 戻　値    :True＝ファイルが存在する、False＝ファイルが存在しない
'************************************************************************************
Public Function existsFile( _
      ByVal path As String _
    , Optional ByRef fso As FileSystemObject = Nothing _
) As Boolean

    Dim wkFso As FileSystemObject

    On Error GoTo ErrorHandler

    'パラメータにFileSystemObjectが渡された場合、それを使用する
    If fso Is Nothing Then
        Set wkFso = New FileSystemObject
    Else
        Set wkFso = fso
    End If

    If wkFso.FileExists(path) Then
        existsFile = True
    Else
        existsFile = False
    End If

    GoTo Fin

ErrorHandler:
    'エラー処理

Fin:
    '後処理
    Set wkFso = Nothing

    'エラーが発生した場合は呼出元に投げる
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If

End Function

'************************************************************************************
' 機　能    :フォルダ存在チェック
' 引　数    :in      path                       フォルダパス
'           :in      fso                        FileSystemObject
' 戻　値    :True＝ファイルが存在する、False＝ファイルが存在しない
'************************************************************************************
Public Function existsFolder( _
      ByVal path As String _
    , Optional ByRef fso As FileSystemObject = Nothing _
) As Boolean

    Dim wkFso As FileSystemObject

    On Error GoTo ErrorHandler

    'パラメータにFileSystemObjectが渡された場合、それを使用する
    If fso Is Nothing Then
        Set wkFso = New FileSystemObject
    Else
        Set wkFso = fso
    End If

    If wkFso.FolderExists(path) Then
        existsFolder = True
    Else
        existsFolder = False
    End If

    GoTo Fin

ErrorHandler:
    'エラー処理

Fin:
    '後処理
    Set wkFso = Nothing

    'エラーが発生した場合は呼出元に投げる
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If

End Function

'************************************************************************************
' 機　能    :拡張子チェック
' 引　数    :in  path                            ファイルパス
'           :in  extensions                      チェックする拡張子（複数の場合はカンマ区切り）
' 戻　値    :True/拡張子が一致する  :  False/一致しない
'************************************************************************************
Public Function checkExtensionName(ByVal path As String, ByVal extensions As String) As Boolean

    Const EXTENSION_DELIMITER As String = ","

    'ファイルの拡張子取得
    Dim val As String
    val = getExtensionName(path)

    'チェックする拡張子を配列に取得
    Dim aryExtenstions As Variant
    aryExtenstions = Split(extensions, EXTENSION_DELIMITER)

    '配列にファイルの拡張子が存在するか
    Dim i As Long
    For i = LBound(aryExtenstions) To UBound(aryExtenstions)
        If aryExtenstions(i) = val Then
            '存在する場合：true
            checkExtensionName = True
            Exit Function
        End If
    Next

    '存在しない場合：false
    checkExtensionName = False

End Function

Public Sub imageTest()

    Dim shp As Shape
    
    For Each shp In shMsg.Shapes
        
        Debug.Print shp.Name

    Next



End Sub
