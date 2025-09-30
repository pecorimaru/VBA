Attribute VB_Name = "modExclEdit"
Option Explicit

'************************************************************************************
' �@  �\    :Excel����̔ėp���W���[��
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/11/01  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************

'************************************************************************************
' �@�@�\    :�ő�s���擾
' ���@��    :in  sh                              �ΏۃV�[�g
'           :in  colIdx                          �ŏI�s�𔻒肷����Index�i���邢�͗�L���j
' �߁@�l    :�ő�s
'************************************************************************************
Public Function getRowMax(ByRef sh As Worksheet, ByVal colIdx As Variant) As Long
    getRowMax = sh.Cells(sh.Rows.count, colIdx).End(xlUp).row
End Function

'************************************************************************************
' �@�@�\    :�ő����擾
' ���@��    :in  sh                              �ΏۃV�[�g
'           :in  rowIdx                          �ŏI��𔻒肷��s��Index
' �߁@�l    :�ő��
'************************************************************************************
Public Function getColMax(ByRef sh As Worksheet, ByVal rowIdx As Long) As Long
    getColMax = sh.Cells(rowIdx, sh.Columns.count).End(xlToLeft).Column
End Function

'************************************************************************************
' �@�@�\    :��L�������ԍ����擾
' ���@��    :in  Alphabet                        ��L��
' �߁@�l    :��ԍ�
'************************************************************************************
Public Function getColFrAlphabet(ByVal Alphabet As String) As Long
    On Error GoTo ARGS_ERROR
    getColFrAlphabet = ActiveSheet.Range(Alphabet & "1").Column
    Exit Function
ARGS_ERROR:
    getColFrAlphabet = 0
End Function

'************************************************************************************
' �@�@�\    :�u�b�N�̃I�[�v���`�F�b�N
' ���@��    :in  xlApp                           �ΏۃA�v���P�[�V����
'           :in  bkNm                            �����u�b�N��
'           :out bkRecv                          �J����Ă���ꍇ�ɃZ�b�g
' �߁@�l    :True/�u�b�N���J����Ă���  :  False/�u�b�N���J����Ă��Ȃ�
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
' �@�@�\    :�u�b�N�̃N���[�Y�i�J���Ă���ꍇ�̂݁j
' ���@��    :in  xlApp                            Excel.Application
'           :in  bkNm                             �u�b�N��
' �߁@�l    :�Ȃ�
'************************************************************************************
Public Sub bkClose(ByRef xlApp As Excel.Application, ByVal bkNm As String)
    If isBkOpen(xlApp, bkNm) Then
        xlApp.Workbooks(bkNm).Close (False)
    End If
End Sub

'************************************************************************************
' �@�@�\    :�V�[�g�̑��݃`�F�b�N
' ���@��    :in  bk                               �Ώۃu�b�N
'           :in  shNm                             �����V�[�g��
'           :out shRecv                           �擾�V�[�g
' �߁@�l    :True/���݂���  :  False/���݂��Ȃ�
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
' �@�@�\    :�uSheet1�v�ȊO�̃V�[�g��S�č폜
' ���@��    :in  tgtBook                         �Ώۃu�b�N
' �߁@�l    :�Ȃ�
'************************************************************************************
Public Sub shClean(ByRef tgtBook As Workbook)

    Dim ws As Worksheet
    
    '�擪�V�[�g�ȊO��S�č폜
    For Each ws In tgtBook.Worksheets
        If ws.Name <> "Sheet1" Then
            ws.Delete
        End If
    Next

    Set ws = Nothing

End Sub

'************************************************************************************
' �@�@�\    :�Z���͈͂�񎟌��z��ɕϊ�
' ���@��    :in  sh                              �V�[�g
'           :in  rowBgn                          �J�n�s
'           :in  rowEnd                          �ŏI�s
'           :in  colBgn                          �J�n��
'           :in  colEnd                          �ŏI��
' �߁@�l    :�Z���͈́i�񎟌��z��j
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
' �@�@�\    :�uSheet1�v�ȊO�̃V�[�g��S�č폜
' ���@��    :in  tgtBook                         �Ώۃu�b�N
' �߁@�l    :�Ȃ�
'************************************************************************************
Public Sub shapesSaveAsPicture(ByRef shp As Shape, ByVal savePath As String)

    '�`���[�g���쐬
    Dim Cht
    Set Cht = ActiveSheet.ChartObjects.add(0, 0, shp.Width, shp.Height)
    
    With Cht
        shp.CopyPicture Format:=xlBitmap '�I�[�g�V�F�C�v���摜�Ƃ��ăR�s�[
        .Chart.Parent.Select '�`���[�g��I��
        .Chart.Paste '�`���[�g�ɓ\��t��
        .Chart.Export savePath '�`���[�g���A�摜�Ƃ��ĕۑ�
        .Delete '�`���[�g���폜
    End With

End Sub

'************************************************************************************
' �@�@�\    :�Ԙg���쐬
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
