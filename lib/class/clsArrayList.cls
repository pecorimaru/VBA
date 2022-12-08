VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'************************************************************************************
' 機　能    :ArrayListクラス
' 説  明    :要素番号が１から始まる動的な配列クラス
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************

'配列
Private pItem As Variant

'************************************************************************************
' 機　能    :デストラクタ
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Private Sub Class_Terminate()
    If Not IsEmpty(pItem) Then
        Erase pItem
    End If
End Sub

'************************************************************************************
' 機　能    :セッター（データ型）
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Property Let item(ByVal aryNum As Long, ByVal elem As Variant)
    pItem(aryNum) = elem
End Property

'************************************************************************************
' 機　能    :セッター（オブジェクト型）
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Property Set item(ByVal aryNum As Long, ByRef elem As Variant)
    Set pItem(aryNum) = elem
End Property

'************************************************************************************
' 機　能    :ゲッター
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Property Get item(ByVal aryNum As Long) As Variant
Attribute item.VB_UserMemId = 0
    If IsObject(pItem(aryNum)) Then
        Set item = pItem(aryNum)
    Else
        item = pItem(aryNum)
    End If
End Property

'************************************************************************************
' 機　能    :要素を追加
' 引  数    :in  elem                            要素
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Sub add(ByRef elem As Variant)

    '配列が未割当ての場合
    If Me.IsAryEmpty Then
        
        '配列を再定義
        ReDim pItem(1 To 1) As Variant

    Else
        
        '配列の内容を保持したまま再定義
        ReDim Preserve pItem(1 To UBound(pItem) + 1) As Variant

    End If

    '追加する要素のチェック
    If IsObject(elem) Then
        
        'オブジェクト型である場合
        Set pItem(UBound(pItem)) = elem

    Else
    
        '基本型である場合
        pItem(UBound(pItem)) = elem

    End If

End Sub

'************************************************************************************
' 機　能    :全ての要素を追加
' 引  数    :in  ary                             配列
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Sub addAll(ByRef ary As clsArrayList)

    Dim iter As clsIterator

    '追加する要素が空である場合
    If ary.IsAryEmpty Then
        Exit Sub
    End If

    '配列が未割当ての場合
    If Me.IsAryEmpty Then
        
        '配列を再定義
        ReDim pItem(1 To 1) As Variant

    End If
        
    '配列の要素を全て追加
    Set iter = ary.iterator
    While iter.hasNext
        Me.add iter.toNext
    Wend
        
    Set iter = Nothing

End Sub

'************************************************************************************
' 機　能    :要素数を取得
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Function count() As Long

    '要素が存在する場合
    If Not Me.IsAryEmpty Then
        '最大の要素番号を返却
        count = UBound(pItem)
    End If

End Function

'************************************************************************************
' 機　能    :配列が存在するか
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Function IsAryEmpty() As Boolean
    IsAryEmpty = IsEmpty(pItem)
End Function

'************************************************************************************
' 機　能    :指定した要素を削除
' 引  数    :in  aryNum                          要素番号
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Sub remove(ByVal aryNum As Long)

    '指定した要素番号から順に要素を詰めて再配置
    Dim i As Long
    For i = aryNum To UBound(pItem) - 1
        item(i) = item(i + 1)
    Next

    '配列を削減した要素数で再定義
    ReDim Preserve pItem(1 To UBound(pItem) - 1) As Variant

End Sub

'************************************************************************************
' 機　能    :要素を全て削除
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Sub removeAll()
    If Not Me.IsAryEmpty Then
        Erase pItem
    End If
End Sub

'************************************************************************************
' 機　能    :指定された要素が配列内で最初に一致する要素番号を返却
' 引  数    :in  idxBgn                          検索開始位置
'           :in  elem                            検索値
'           :in  isPart                          部分一致検索であるか
'           :in  rtnIdx                          何件目の一致を返却するか
' 戻  値    :最初に一致した要素番号
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Function indexOf( _
      ByVal elem As Variant _
    , Optional ByVal idxBgn As Long = 1 _
    , Optional ByVal isPart As Boolean = False _
    , Optional ByVal rtnIdx As Long = 1 _
) As Long

    Dim isMatch As Boolean

    '戻り値をクリア
    indexOf = 0

    On Error GoTo ErrorHandler

    '配列が空である場合
    If Me.IsAryEmpty Then
        GoTo Fin
    End If

    Dim hitCnt As Long
    hitCnt = 0

    '先頭の要素から順に最後の要素まで走査
    Dim aryIdx As Long
    For aryIdx = idxBgn To UBound(pItem)
    
        If isPart Then
            '部分一致する場合
            isMatch = pItem(aryIdx) Like "*" & elem & "*"
        Else
            '完全一致する場合
            isMatch = pItem(aryIdx) = elem
        End If
        
        If isMatch Then

            hitCnt = hitCnt + 1
        
            If hitCnt = rtnIdx Then

                '要素番号を返却
                indexOf = aryIdx
                Exit Function
            
            End If
            
        End If

    Next

    '要素が存在しない
    indexOf = 0
    GoTo Fin

ErrorHandler:
    '************************************************************
    ' エラー処理
    '************************************************************
    
    '要素が存在しない
    indexOf = 0
    
    Resume Fin
Fin:

End Function

'************************************************************************************
' 機　能    :指定された要素が配列内で最後に一致する要素番号を返却
' 引  数    :in  idxBgn                          検索開始位置
'           :in  elem                            検索値
'           :in  isPart                          部分一致検索であるか
'           :in  rtnIdx                          何件目の一致を返却するか
' 戻  値    :最初に一致した要素番号
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Function lastIndexOf( _
      ByVal elem As Variant _
    , Optional ByVal idxBgn As Long = 1 _
    , Optional ByVal isPart As Boolean = False _
    , Optional ByVal rtnIdx As Long = 1 _
) As Long

    Dim isMatch As Boolean

    '戻り値をクリア
    lastIndexOf = 0

    On Error GoTo ErrorHandler

    '配列が空である場合
    If Me.IsAryEmpty Then
        GoTo Fin
    End If

    Dim hitCnt As Long
    hitCnt = 0

    '先頭の要素から順に最後の要素まで走査
    Dim aryIdx As Long
    For aryIdx = idxBgn To LBound(pItem) Step -1
    
        If isPart Then
            '部分一致する場合
            isMatch = pItem(aryIdx) Like "*" & elem & "*"
        Else
            '完全一致する場合
            isMatch = pItem(aryIdx) = elem
        End If
        
        If isMatch Then
        
            hitCnt = hitCnt + 1
        
            If hitCnt = rtnIdx Then
        
                '要素番号を返却
                lastIndexOf = aryIdx
                Exit Function
            
            End If
            
        End If

    Next

    '要素が存在しない
    lastIndexOf = 0
    GoTo Fin

ErrorHandler:
    '************************************************************
    ' エラー処理
    '************************************************************
    
    '要素が存在しない
    lastIndexOf = 0
    
    Resume Fin
Fin:

End Function

'************************************************************************************
' 機　能    :指定された要素が配列内に存在するか
' 引  数    :in  elem                            検索値
' 戻  値    :True/存在する  ：  False/存在しない
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Function contains(ByVal elem As Variant) As Boolean

    '戻り値をクリア
    contains = False

    '要素の存在チェック
    If Me.indexOf(elem) <> 0 Then
        contains = True
    End If

End Function

'************************************************************************************
' 機　能    :イテレータを取得
' -----------------------------------------------------------------------------------
' 履　歴　　:2020/04/10  J.Tamura  Ver.1.0.00  新規作成
'************************************************************************************
Public Function iterator() As clsIterator

    Dim iter As New clsIterator

    '配列が空である場合
    If Me.IsAryEmpty Then
        Set iterator = iter
        GoTo Fin
    End If

    'イテレータに配列を設定
    iter.setAry = pItem

    '設定したイテレータを返却
    Set iterator = iter
    
Fin:

    Set iter = Nothing

End Function
