Attribute VB_Name = "modStringUtils"
Option Explicit

'************************************************************************************
' 機　能    :文字列操作の汎用モジュール
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/11/01  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************

'************************************************************************************
' 機　能    :未入力チェック
' 引　数    :in  val                             判定文字列
' 戻　値    :True/未入力である  :  False/未入力ではない
'************************************************************************************
Public Function isBlank(ByVal str As String) As Boolean
    isBlank = False
    If str <> "" And Not IsNull(str) Then
        Exit Function
    End If
    isBlank = True
End Function

'************************************************************************************
' 機　能    :長さチェック
' 引　数    :in  val                             判定文字列
' 戻　値    :True/長さがある  :  False/長さがない
'************************************************************************************
Public Function hasLength(ByVal str As String) As Boolean
    hasLength = False
    If Len(str) = 0 Then
        Exit Function
    End If
    hasLength = True
End Function

'************************************************************************************
' 機　能    :プレースホルダの置換
' 引　数    :in  expression                       置換元
'           :in  replaces()                       置換文字
' 戻　値    :置換後文字列
'************************************************************************************
Public Function ReplacePh(ByVal expression As String, ParamArray replaces()) As String
    Dim edit As String
    edit = expression
    Dim num As Integer
    num = 0
    Dim rep As Variant
    For Each rep In replaces
        num = num + 1
        edit = Replace(edit, "{" & CStr(num) & "}", CStr(rep))
    Next
    ReplacePh = edit
End Function
