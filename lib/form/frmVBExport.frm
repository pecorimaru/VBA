VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVBExport 
   Caption         =   "VBA���W���[���̃G�N�X�|�[�g"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   105
   ClientWidth     =   8565.001
   OleObjectBlob   =   "frmVBExport.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmVBExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************************
' �@�@�\    :VBA���W���[���̃G�N�X�|�[�g
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/10/18  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************

'************************************************************************************
' �@�@�\    :���[�U�[�t�H�[����������
' ���@��    :�J���Ă���u�b�N����S�Ď擾���A[���[�N�u�b�N�I��]�ɃZ�b�g
' ���@��    :�Ȃ�
' �߁@�l    :�Ȃ�
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/10/18  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Private Sub UserForm_Initialize()
    Dim bk As Variant
    For Each bk In Application.Workbooks
        With frmVBExport.cmbBkExpt
            .AddItem bk.Name
        End With
    Next
End Sub

'************************************************************************************
' �@�@�\    :�Q�ƃ{�^���i�t�@�C���Q�Ɓj����������
' ���@��    :�t�@�C���I���_�C�A���O���J���A�I�������t�@�C�������[�N�u�b�N�ɃZ�b�g
' ���@��    :�Ȃ�
' �߁@�l    :�Ȃ�
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/10/18  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Private Sub btnRefBkExpt_Click()
    With frmVBExport.txtBkExpt
        .Value = showFilePickerDialog("�u�b�N��I��", .Value, "*.xlsm;*.xlam")
    End With
End Sub

'************************************************************************************
' �@�@�\    :�Q�ƃ{�^���i�ۑ���t�H���_�j����������
' ���@��    :�t�H���_�I���_�C�A���O���J���A�I�������t�H���_��ۑ���t�H���_�ɃZ�b�g
' ���@��    :�Ȃ�
' �߁@�l    :�Ȃ�
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/10/18  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Private Sub btnRefSavePath_Click()
    With frmVBExport.txtSavePath
        .Value = showFolderPickerDialog("�ۑ����I��", .Value)
    End With
End Sub

'************************************************************************************
' �@�@�\    :�I�v�V�����{�^���i���݊J���Ă���u�b�N�j����������
' ���@��    :�Ȃ�
' �߁@�l    :�Ȃ�
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/10/18  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Private Sub optOpening_Click()
    frmVBExport.cmbBkExpt.Enabled = True
    frmVBExport.txtBkExpt.Enabled = False
    frmVBExport.btnRefBkExpt.Enabled = False
End Sub

'************************************************************************************
' �@�@�\    :�I�v�V�����{�^���i���݊J���Ă���u�b�N�j����������
' ���@��    :�Ȃ�
' �߁@�l    :�Ȃ�
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/10/18  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Private Sub optRefFile_Click()
    frmVBExport.cmbBkExpt.Enabled = False
    frmVBExport.txtBkExpt.Enabled = True
    frmVBExport.btnRefBkExpt.Enabled = True
End Sub

'************************************************************************************
' �@�@�\    :�G�N�X�|�[�g�{�^������������
' ���@��    :�Ȃ�
' �߁@�l    :�Ȃ�
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/10/18  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Private Sub btnExecExpt_Click()

    Dim module As VBComponent
    Dim moduleList As VBComponents
    Dim bkExpt As Workbook

    On Error GoTo ErrorHandler

    '************************************************************
    ' ��������
    '************************************************************

    '�������s�t���O��������
    Dim isJobed As Boolean
    isJobed = False

    '���̓`�F�b�N
    If Not isValidInput() Then
        GoTo Fin
    End If

    '�G�N�X�|�[�g�Ώۃu�b�N���擾
    If frmVBExport.optOpening = True Then
        Set bkExpt = Workbooks(frmVBExport.cmbBkExpt.Value)
    Else
        Set bkExpt = Workbooks.Open(frmVBExport.txtBkExpt.Value)
    End If
    
    '�u�b�N�̃��W���[���ꗗ���擾
    Set moduleList = bkExpt.VBProject.VBComponents
    
    '************************************************************
    ' ���C������
    '************************************************************
    
    'VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList

        Dim isExpt As Boolean          '�G�N�X�|�[�g�Ώۃt���O
        Dim ext As String              '�g���q
        
        '���W���[�����ƂɑΏ۔���Ɗg���q�𕪊�
        Select Case module.Type
            
            '�N���X���W���[��
            Case vbext_ct_ClassModule

                isExpt = frmVBExport.chkClsMod.Value
                ext = "cls"

            '���[�U�[�t�H�[��
            Case vbext_ct_MSForm
                
                isExpt = frmVBExport.chkUserForm.Value
                ext = "frm"

            '�W�����W���[��
            Case vbext_ct_StdModule

                isExpt = frmVBExport.chkStdMod.Value
                ext = "bas"
                
            '��L�ȊO
            Case Else
            
                isExpt = frmVBExport.chkBkSh.Value
                ext = "cls"
        
        End Select

        '�Ώۃ��W���[���Ƀ`�F�b�N������ꍇ�A�G�N�X�|�[�g
        If isExpt Then
            Dim exptPath As String
            exptPath = frmVBExport.txtSavePath.Value & "\" & module.Name & "." & ext
        
            '�㏑���`�F�b�N
            If existsFile(exptPath) Then
                Dim rc As Integer
                rc = showMsgBox(getMsg("MQ0002", exptPath), xlInfo, vbYesNo)
                If rc = 1 Then
                    Call module.Export(exptPath)
                    isJobed = True
                End If
            Else
                Call module.Export(exptPath)
                isJobed = True
            End If
        End If
    
    Next

    '�u�b�N���N���[�Y
    If frmVBExport.optOpening = False Then
        bkExpt.Close
    End If

    '����I��
    If isJobed Then
        Call showMsgBox(getMsg("MI0002", "�G�N�X�|�[�g"), xlInfo)
    Else
        Call showMsgBox(getMsg("MI0003", "�G�N�X�|�[�g"), xlInfo)
    End If
    GoTo Fin

ErrorHandler:
    '************************************************************
    ' �G���[����
    '************************************************************

    '�ُ�I��
    Call showSystemErrorMsg(Err, ThisWorkbook.Name)
    Resume Fin

Fin:
    '************************************************************
    ' �㏈��
    '************************************************************

    Set bkExpt = Nothing
    Set moduleList = Nothing
    Set module = Nothing

End Sub

'************************************************************************************
' �@�@�\    :���̓`�F�b�N
' ���@��    :�Ȃ�
' �߁@�l    :�Ȃ�
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/10/18  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Private Function isValidInput() As Boolean

    '�߂�l���N���A
    isValidInput = False

    On Error GoTo ErrorHandler

    '************************************************************
    ' ���C������
    '************************************************************

    With frmVBExport

    '----���[�N�u�b�N�I��----

        If frmVBExport.optOpening = True Then

            If isBlank(.cmbBkExpt.Value) Then
                Call showMsgBox(getMsg("ME0002", "���[�N�u�b�N"), xlWarning)
                isValidInput = False
                GoTo Fin
            End If
    
        Else
    
            If isBlank(.txtBkExpt.Value) Then
                Call showMsgBox(getMsg("ME0001", "���[�N�u�b�N"), xlWarning)
                isValidInput = False
                GoTo Fin
            End If
    
        End If
    
    '----�ۑ���t�H���_----
    
        If isBlank(.txtSavePath.Value) Then
            Call showMsgBox(getMsg("ME0001", "�ۑ���t�H���_"), xlWarning)
            isValidInput = False
            GoTo Fin
        End If
    
        If Not existsFolder(.txtSavePath.Value) Then
            Call showMsgBox(getMsg("ME0003", "�ۑ���t�H���_", .txtSavePath.Value), xlWarning)
            isValidInput = False
            GoTo Fin
        End If

    End With

    '����I��
    isValidInput = True
    GoTo Fin

ErrorHandler:
    '************************************************************
    ' �G���[����
    '************************************************************

    '�ُ�I��
    Call showSystemErrorMsg(Err, ThisWorkbook.Name)
    isValidInput = False
    Resume Fin

Fin:
    '************************************************************
    ' �㏈��
    '************************************************************

End Function
