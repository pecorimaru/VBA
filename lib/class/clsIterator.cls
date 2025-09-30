VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'************************************************************************************
' �@�@�\    :Iterator�N���X
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' ���@���@�pecorimaru4/10  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************

'�ݒ�z��
Private pIter As Variant

'�v�f�ԍ��i�J�[�\���ʒu�j
Private pNumCursor As Long

'�v�f�ԍ��i�o�C�A�X�j
Private pNumBias As Long

'************************************************************************************
' �@�@�\    :�R���X�g���N�^
' -----------------------------------------------------------------------------------
' ���@���@�pecorimaru4/10  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Private Sub Class_Initialize()
    pNumCursor = 0
    pNumBias = 0
End Sub
'************************************************************************************
' �@�@�\    :�f�X�g���N�^
' -----------------------------------------------------------------------------------
' ���@���@�pecorimaru4/10  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Private Sub Class_Terminate()
    If Not IsEmpty(pIter) Then
        Erase pIter
    End If
    pNumCursor = 0
    pNumBias = 0
End Sub

'************************************************************************************
' �@�@�\    :�Z�b�^�[
' -----------------------------------------------------------------------------------
' ���@���@�pecorimaru4/10  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Public Property Let setAry(ByVal elem As Variant)
    
    '�z����擾
    pIter = elem
    
    '�ŏ��v�f�ԍ����擾
    pNumBias = LBound(pIter)

    '�J�[�\���ʒu���ŏ��v�f�ԍ��ŏ�����
    pNumCursor = pNumBias - 1

End Property

'************************************************************************************
' �@�@�\    :���̗v�f�����݂��邩
' ��  ��    :�v�f�ԍ���v�f���Ŋ������]��̒l����A���̗v�f�����݂��邩���f����
' -----------------------------------------------------------------------------------
' ���@���@�pecorimaru4/10  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Public Function hasNext() As Boolean

    '�߂�l���N���A
    hasNext = False

    '�z�񂪋�ł���ꍇ
    If IsEmpty(pIter) Then
        Exit Function
    End If

    '���̗v�f�ԍ����ő�v�f�ԍ��Ŋ������]�肪0�ł���ꍇ
    If (pNumCursor + 1) Mod (UBound(pIter) + 1) = 0 Then

        '���̗v�f�����݂��Ȃ�
        hasNext = False
        Exit Function

    End If
    
    '���̗v�f�����݂���
    hasNext = True

End Function

'************************************************************************************
' �@�@�\    :���̗v�f��ݒ肷��
' -----------------------------------------------------------------------------------
' ���@���@�pecorimaru4/10  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************
Public Function toNext() As Variant

    '���̗v�f�����݂���ꍇ�͎��̗v�f��ݒ�
    '���̗v�f�����݂��Ȃ��ꍇ�͍ŏ��̗v�f��ݒ�
    pNumCursor = ((pNumCursor + 1 - pNumBias) Mod (UBound(pIter) + 1 - pNumBias)) + pNumBias

    '���̗v�f��ݒ�
    If IsObject(pIter(pNumCursor)) Then
        
        '�I�u�W�F�N�g�^�ł���ꍇ
        Set toNext = pIter(pNumCursor)
    Else
    
        '�f�[�^�^�ł���ꍇ
        toNext = pIter(pNumCursor)
    End If

End Function
