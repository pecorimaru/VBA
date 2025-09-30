Attribute VB_Name = "modFileUtils"
Option Explicit

'************************************************************************************
' �@�@�\    :�t�@�C������̔ėp���W���[��
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/11/01  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************

'************************************************************************************
' �@�@�\    :�t�@�C���I���_�C�A���O�i�ėp�F�t�B���^�P��ށ^�P���I���j
' ���@��    :in  ttl                             �_�C�A���O�̃^�C�g��
'           :in  filterName                      �t�B���^��
'           :in  filterExtension                 �t�B���^�g���q
'           :in  initialFileName                 �����I��l
' �߁@�l    :�I���t�@�C����
'************************************************************************************
Public Function showFilePickerDialog( _
      ByVal ttl As String _
    , ByVal filterName As String _
    , ByVal filterExtension As String _
    , Optional ByVal initialFileName As String = "" _
) As String

    '�߂�l���N���A
    showFilePickerDialog = ""

    '�_�C�A���O���
    With Application.FileDialog(msoFileDialogOpen)

        '�_�C�A���O�̃^�C�g��
        .title = ttl

        '�_�C�A���O�̃t�B���^
        .Filters.Clear
        .Filters.add filterName, filterExtension
        .FilterIndex = 1

        '�����t�@�C���I���������Ȃ�
        .AllowMultiSelect = False

        '�_�C�A���O�̏����I��l
        If Trim(initialFileName) <> "" Then
            .initialFileName = initialFileName
        End If

        '�_�C�A���O��\��
        If .Show <> 0 Then
            If .SelectedItems.count > 0 Then
                '�t�@�C�����I�����ꂽ�ꍇ�F�߂�l�ɃZ�b�g
                showFilePickerDialog = .SelectedItems.item(1)
            End If
        End If
    
    End With

End Function

'************************************************************************************
' �@�@�\    :�t�@�C���I���_�C�A���O�i�ėp�F�t�B���^�P��ށ^�����I���\�j
' ���@��    :in  ttl                             �_�C�A���O�̃^�C�g��
'           :in  filterName                      �t�B���^��
'           :in  filterExtension                 �t�B���^�g���q
'           :in  initialFileName                 �����I��l
' �߁@�l    :�I���t�@�C���p�X���X�g
'************************************************************************************
Public Function showFilesPickerDialog( _
      ByVal ttl As String _
    , ByVal filterName As String _
    , ByVal filterExtension As String _
    , Optional ByVal initialFileName As String = "" _
) As String()

    Dim selItemList() As String

    '�_�C�A���O���
    With Application.FileDialog(msoFileDialogOpen)

        '�_�C�A���O�̃^�C�g��
        .title = ttl

        '�_�C�A���O�̃t�B���^
        .Filters.Clear
        .Filters.add filterName, filterExtension
        .FilterIndex = 1

        '�����I��������
        .AllowMultiSelect = True

        '�_�C�A���O�̏����I��l
        If Trim(initialFileName) <> "" Then
            .initialFileName = initialFileName
        End If

        '�_�C�A���O��\��
        If .Show <> 0 Then
            If .SelectedItems.count > 0 Then
                
                '�I���t�@�C�����X�g
                ReDim selItemList(0) As String
                
                '�I�������t�@�C���������[�v
                Dim selItem As Variant
                For Each selItem In .SelectedItems
                                        
                    '�z��̗v�f�����݂���ꍇ
                    If selItemList(0) <> "" Then
                     
                         '�z����g��
                        ReDim Preserve selItemList(UBound(selItemList) + 1)
                     
                    End If
                                         
                    '�I���t�@�C����v�f�ɒǉ�
                    selItemList(UBound(selItemList)) = selItem
                     
                 Next selItem
            
            End If
        End If
    
    End With

    '�I���t�@�C���p�X���X�g��ԋp
    showFilesPickerDialog = selItemList

End Function

'************************************************************************************
' �@�@�\    :�t�H���_�I���_�C�A���O�i�ėp�j
' ���@��    :in  title                           �_�C�A���O�̃^�C�g��
'           :in  initialFolderPath               �t�B���^��
' �߁@�l    :�I���t�H���_��
'************************************************************************************
Public Function showFolderPickerDialog( _
      ByVal title As String _
    , Optional ByVal initialFolderPath As String = "" _
) As String
    
    '�߂�l���N���A
    showFolderPickerDialog = ""

    '�t�H���_�I���_�C�A���O�ݒ�
    With Application.FileDialog(msoFileDialogFolderPicker)
        
        '�_�C�A���O�̃^�C�g��
        .title = title
        
        '�_�C�A���O�̏����I��l
        If Trim(initialFolderPath) <> "" Then
            .initialFileName = initialFolderPath
        End If

        '�_�C�A���O��\��
        If .Show <> 0 Then
            If .SelectedItems.count > 0 Then
                '�t�H���_���I�����ꂽ�ꍇ�F�߂�l�ɃZ�b�g
                showFolderPickerDialog = .SelectedItems.item(1)
            End If
        End If
    
    End With

End Function

'************************************************************************************
' �@�@�\    :�t�@�C�������擾(�g���q���܂�)
' ���@��    :in      path                       �t�@�C���p�X
' �߁@�l    :�t�@�C����(�g���q���܂�)
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
' �@�@�\    :�t�@�C�������擾(�g���q���܂܂Ȃ�)
' ���@��    :in      path                       �t�@�C���p�X
' �߁@�l    :�t�@�C����(�g���q���܂܂Ȃ�)
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
' �@�@�\    :�g���q���擾
' ���@��    :in      path                       �t�@�C���p�X
' �߁@�l    :�g���q
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
' �@�@�\    :�t�H���_�p�X���擾
' ���@��    :in      path                       �t�@�C���p�X
' �߁@�l    :�t�H���_���p�X
'************************************************************************************
Public Function getFolderFromPath(ByVal path As String) As String

    Dim wkPath As String
    wkPath = path

    '������"\"�̏ꍇ�A�폜
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
' �@�@�\    :�t�@�C�����݃`�F�b�N
' ���@��    :in      path                       �t�@�C���p�X
'           :in      fso                        FileSystemObject
' �߁@�l    :True���t�@�C�������݂���AFalse���t�@�C�������݂��Ȃ�
'************************************************************************************
Public Function existsFile( _
      ByVal path As String _
    , Optional ByRef fso As FileSystemObject = Nothing _
) As Boolean

    Dim wkFso As FileSystemObject

    On Error GoTo ErrorHandler

    '�p�����[�^��FileSystemObject���n���ꂽ�ꍇ�A������g�p����
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
    '�G���[����

Fin:
    '�㏈��
    Set wkFso = Nothing

    '�G���[�����������ꍇ�͌ďo���ɓ�����
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If

End Function

'************************************************************************************
' �@�@�\    :�t�H���_���݃`�F�b�N
' ���@��    :in      path                       �t�H���_�p�X
'           :in      fso                        FileSystemObject
' �߁@�l    :True���t�@�C�������݂���AFalse���t�@�C�������݂��Ȃ�
'************************************************************************************
Public Function existsFolder( _
      ByVal path As String _
    , Optional ByRef fso As FileSystemObject = Nothing _
) As Boolean

    Dim wkFso As FileSystemObject

    On Error GoTo ErrorHandler

    '�p�����[�^��FileSystemObject���n���ꂽ�ꍇ�A������g�p����
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
    '�G���[����

Fin:
    '�㏈��
    Set wkFso = Nothing

    '�G���[�����������ꍇ�͌ďo���ɓ�����
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If

End Function

'************************************************************************************
' �@�@�\    :�g���q�`�F�b�N
' ���@��    :in  path                            �t�@�C���p�X
'           :in  extensions                      �`�F�b�N����g���q�i�����̏ꍇ�̓J���}��؂�j
' �߁@�l    :True/�g���q����v����  :  False/��v���Ȃ�
'************************************************************************************
Public Function checkExtensionName(ByVal path As String, ByVal extensions As String) As Boolean

    Const EXTENSION_DELIMITER As String = ","

    '�t�@�C���̊g���q�擾
    Dim val As String
    val = getExtensionName(path)

    '�`�F�b�N����g���q��z��Ɏ擾
    Dim aryExtenstions As Variant
    aryExtenstions = Split(extensions, EXTENSION_DELIMITER)

    '�z��Ƀt�@�C���̊g���q�����݂��邩
    Dim i As Long
    For i = LBound(aryExtenstions) To UBound(aryExtenstions)
        If aryExtenstions(i) = val Then
            '���݂���ꍇ�Ftrue
            checkExtensionName = True
            Exit Function
        End If
    Next

    '���݂��Ȃ��ꍇ�Ffalse
    checkExtensionName = False

End Function

Public Sub imageTest()

    Dim shp As Shape
    
    For Each shp In shMsg.Shapes
        
        Debug.Print shp.Name

    Next



End Sub
