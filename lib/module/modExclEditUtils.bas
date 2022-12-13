Attribute VB_Name = "modExclEdit"
Option Explicit

'************************************************************************************
' 機  能    :Excel操作の汎用モジュール
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/11/01  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************

'************************************************************************************
' 機　能    :最大行を取得
' 引　数    :in  sh                              対象シート
'           :in  colIdx                          最終行を判定する列のIndex（あるいは列記号）
' 戻　値    :最大行
'************************************************************************************
Public Function getRowMax(ByRef sh As Worksheet, ByVal colIdx As Variant) As Long
    getRowMax = sh.Cells(sh.Rows.count, colIdx).End(xlUp).row
End Function

'************************************************************************************
' 機　能    :最大列を取得
' 引　数    :in  sh                              対象シート
'           :in  rowIdx                          最終列を判定する行のIndex
' 戻　値    :最大列
'************************************************************************************
Public Function getColMax(ByRef sh As Worksheet, ByVal rowIdx As Long) As Long
    getColMax = sh.Cells(rowIdx, sh.Columns.count).End(xlToLeft).Column
End Function

'************************************************************************************
' 機　能    :列記号から列番号を取得
' 引　数    :in  Alphabet                        列記号
' 戻　値    :列番号
'************************************************************************************
Public Function getColFrAlphabet(ByVal Alphabet As String) As Long
    On Error GoTo ARGS_ERROR
    getColFrAlphabet = ActiveSheet.Range(Alphabet & "1").Column
    Exit Function
ARGS_ERROR:
    getColFrAlphabet = 0
End Function

'************************************************************************************
' 機　能    :ブックのオープンチェック
' 引　数    :in  xlApp                           対象アプリケーション
'           :in  bkNm                            検索ブック名
'           :out bkRecv                          開かれている場合にセット
' 戻　値    :True/ブックが開かれている  :  False/ブックが開かれていない
'************************************************************************************
Public Function isBkOpen( _
      ByRef xlApp As Excel.Application _
    , ByVal bkNm As String _
    , Optional ByRef bkRecv As Workbook _
) As Boolean
    Dim bk As Workbook
    For Each bk In xlApp.Workbooks
        If bk.Name = bkNm Then
            isBkOpen = True
            Set bkRecv = bk
            Set bk = Nothing
            Exit Function
        End If
    Next
    isBkOpen = False
    Set bk = Nothing
End Function

'************************************************************************************
' 機　能    :ブックのクローズ（開いている場合のみ）
' 引　数    :in  xlApp                            Excel.Application
'           :in  bkNm                             ブック名
' 戻　値    :なし
'************************************************************************************
Public Sub bkClose(ByRef xlApp As Excel.Application, ByVal bkNm As String)
    If isBkOpen(xlApp, bkNm) Then
        xlApp.Workbooks(bkNm).Close (False)
    End If
End Sub

'************************************************************************************
' 機　能    :シートの存在チェック
' 引　数    :in  bk                               対象ブック
'           :in  shNm                             検索シート名
'           :out shRecv                           取得シート
' 戻　値    :True/存在する  :  False/存在しない
'************************************************************************************
Public Function shExists( _
      ByRef bk As Workbook _
    , ByVal shNm As String _
    , Optional ByRef shRecv As Worksheet _
) As Boolean
    Dim sh As Worksheet
    For Each sh In bk.Worksheets
        If sh.Name = shNm Then
            shExists = True
            Set shRecv = sh
            Set sh = Nothing
            Exit Function
        End If
    Next
    shExists = False
    Set sh = Nothing
End Function

'************************************************************************************
' 機　能    :「Sheet1」以外のシートを全て削除
' 引　数    :in  tgtBook                         対象ブック
' 戻　値    :なし
'************************************************************************************
Public Sub shClean(ByRef tgtBook As Workbook)

    Dim ws As Worksheet
    
    '先頭シート以外を全て削除
    For Each ws In tgtBook.Worksheets
        If ws.Name <> "Sheet1" Then
            ws.Delete
        End If
    Next

    Set ws = Nothing

End Sub

'************************************************************************************
' 機　能    :セル範囲を二次元配列に変換
' 引　数    :in  sh                              シート
'           :in  rowBgn                          開始行
'           :in  rowEnd                          最終行
'           :in  colBgn                          開始列
'           :in  colEnd                          最終列
' 戻　値    :セル範囲（二次元配列）
'************************************************************************************
Public Function cnvRange2Ary( _
    ByRef sh As Worksheet _
  , ByVal rowBgn As Long _
  , ByVal rowEnd As Long _
  , ByVal colBgn As Long _
  , ByVal colEnd As Long _
) As Variant
    cnvRange2Ary = sh.Range(sh.Cells(rowBgn, colBgn), sh.Cells(rowEnd, colEnd)).Value
End Function

'************************************************************************************
' 機　能    :「Sheet1」以外のシートを全て削除
' 引　数    :in  tgtBook                         対象ブック
' 戻　値    :なし
'************************************************************************************
Public Sub shapesSaveAsPicture(ByRef shp As Shape, ByVal savePath As String)

    'チャートを作成
    Dim Cht
    Set Cht = ActiveSheet.ChartObjects.add(0, 0, shp.Width, shp.Height)
    
    With Cht
        shp.CopyPicture Format:=xlBitmap 'オートシェイプを画像としてコピー
        .Chart.Parent.Select 'チャートを選択
        .Chart.Paste 'チャートに貼り付け
        .Chart.Export savePath 'チャートを、画像として保存
        .Delete 'チャートを削除
    End With

End Sub

'************************************************************************************
' 機　能    :赤枠を作成
'************************************************************************************
Public Sub mkFrmAttn()

    Dim frmAttn As Shape

    Set frmAttn = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.Left, Selection.Top, 124, 42)

    redFrame.Fill.Visible = msoFalse
    With frmAttn.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
        .Weight = 2
    End With

    frmAttn.Select

    Set frmAttn = Nothing

End Sub
