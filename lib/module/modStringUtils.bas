Attribute VB_Name = "modStringUtils"
Option Explicit

'************************************************************************************
' �@�@�\    :�����񑀍�̔ėp���W���[��
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/11/01  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************

'************************************************************************************
' �@�@�\    :�����̓`�F�b�N
' ���@��    :in  val                             ���蕶����
' �߁@�l    :True/�����͂ł���  :  False/�����͂ł͂Ȃ�
'************************************************************************************
Public Function isBlank(ByVal str As String) As Boolean
    isBlank = False
    If str <> "" And Not IsNull(str) Then
        Exit Function
    End If
    isBlank = True
End Function

'************************************************************************************
' �@�@�\    :�����`�F�b�N
' ���@��    :in  val                             ���蕶����
' �߁@�l    :True/����������  :  False/�������Ȃ�
'************************************************************************************
Public Function hasLength(ByVal str As String) As Boolean
    hasLength = False
    If Len(str) = 0 Then
        Exit Function
    End If
    hasLength = True
End Function

'************************************************************************************
' �@�@�\    :�v���[�X�z���_�̒u��
' ���@��    :in  expression                       �u����
'           :in  replaces()                       �u������
' �߁@�l    :�u���㕶����
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
