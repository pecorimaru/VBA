Attribute VB_Name = "modAppProc"
Option Explicit

'************************************************************************************
' �@�@�\    :�A�v���P�[�V�������샂�W���[��
' �ˑ��֌W  :  shMsg
'           :  modExclEdit
'           :  modFileUtils
'           :  modStringUtils
' Ver       :1.0.00
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/11/01  pecorimaru  Ver.1.0.00  �V�K�쐬
'************************************************************************************

'----------�O���[�o���萔----------

'���b�Z�[�W�{�b�N�X�p
Public Enum XlMsgType
    xlInfo      '�C���t�H���[�V�����A�ʒm
    xlWarning   '���[�j���O�A�x��
    xlError     '�G���[
End Enum

'----------message�V�[�g----------

'�sIndex
Private Enum ROW_SH_MSG
    IDX_HDR = 1
    IDX_DTL_BGN = 2
End Enum

'���b�Z�[�W ��Index
Private Enum COL_SH_MSG
    IDX_ID = 1
    IDX_MSG = 2
End Enum

'----------���W���[���ϐ�----------

'�A�v���P�[�V�����ݒ� �ۑ��p
Private Type AppSettingsType
    enableEventsValue As Boolean
    interactiveValue As Boolean
    screenUpdatingValue As Boolean
    cursorValue As XlMousePointer
    calculationValue As XlCalculation
    displayStatusBarValue As Boolean
    statusBarValue As Variant
End Type

'�A�v���P�[�V�����ݒ� �ۑ��p
Private appSettingsSave As AppSettingsType         '�����J�n�O
Private appSettingsPauseSave As AppSettingsType    '�ꎞ��~�O

'�A�v���P�[�V�����ݒ� �t���O
Private isProcessing As Boolean
Private isPausing As Boolean

'************************************************************************************
' �@�@�\    :���b�Z�[�W�{�b�N�X�\��
' ���@��    :in      prompt                     ���b�Z�[�W���e
'           :in      msgType                    ���b�Z�[�W���
'           :                                     �ExlInfo   (�C���t�H���[�V�����A�ʒm)
'           :                                     �ExlWarning(���[�j���O�A�x��)
'           :                                     �ExlError  (�G���[)
'           :in      buttons                    �{�^����A�C�R�����
'           :in      ttl                        �^�C�g��
'           :in      withAutoCalc               �Čv�Z�t���O
' �߁@�l    :�������ꂽ�{�^���̎��
'************************************************************************************
Public Function showMsgBox( _
      ByVal prompt As String _
    , ByRef msgType As XlMsgType _
    , Optional ByVal buttons As VbMsgBoxStyle = VbMsgBoxStyle.vbOKOnly _
    , Optional ByVal ttl As String = vbNullString _
    , Optional ByVal withAutoCalc As Boolean = True _
) As VbMsgBoxResult

    On Error GoTo ErrorHandler

    '�������ꎞ���f
    Call setApplicationSettingsPause(False)
    
    '�Čv�Z
    If withAutoCalc Then
        Application.Calculate
    End If

    'buttons�p�����[�^����A�C�R���w��𒊏o
    Dim wkButtons As Long
    Dim icon As Long
    wkButtons = buttons
    icon = (wkButtons And &HF0)
    
    '�A�C�R���w�肠��̏ꍇ
    If icon > 0 Then
        '�A�C�R���w����폜
        wkButtons = wkButtons - icon
    End If
    
    '���b�Z�[�W��ʂɍ��킹�ăA�C�R�����w��
    Select Case msgType
        Case XlMsgType.xlInfo
            wkButtons = wkButtons + VbMsgBoxStyle.vbInformation
        Case XlMsgType.xlWarning
            wkButtons = wkButtons + VbMsgBoxStyle.vbExclamation
        Case XlMsgType.xlError
            wkButtons = wkButtons + VbMsgBoxStyle.vbCritical
    End Select
    
    '�^�C�g�����w�肳��Ȃ������ꍇ
    If StrPtr(ttl) = 0 Then
        '���[�N�u�b�N�����^�C�g���Ɏw��
        ttl = getFileNameNoneExtFromPath(ThisWorkbook.Name)
    End If

    showMsgBox = MsgBox(prompt, wkButtons, ttl)

    GoTo Fin

ErrorHandler:
    '�G���[����

Fin:
    '�������ĊJ
    Call setApplicationSettingsReStart

    '�G���[�����������ꍇ�͌ďo���ɓ�����
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If

End Function

'************************************************************************************
' �@�@�\    :�V�X�e���G���[���b�Z�[�W��\������
' ���@��    :in      errObj                     �G���[�I�u�W�F�N�g
'           :in      ttl                        �^�C�g��
' �߁@�l    :�Ȃ�
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
' �@�@�\    :���b�Z�[�W�擾
' ���@��    :in  id                             ���b�Z�[�W�h�c
'           :in  replaceStrings                 �u������������(�����w���)
' �߁@�l    :���b�Z�[�W
'************************************************************************************
Public Function getMsg( _
      ByVal id As String _
    , ParamArray replaceStrings()) As String

    On Error GoTo ErrorHandler

    '���b�Z�[�W�V�[�g�̍ő�s�擾
    Dim rowMax As Long
    rowMax = getRowMax(shMsg, COL_SH_MSG.IDX_ID)
    If rowMax < ROW_SH_MSG.IDX_DTL_BGN Then
        getMsg = ""
        GoTo Fin
    End If

    '���b�Z�[�W�V�[�g�̓��e��񎟌��z��Ɏ擾
    Dim msgAry As Variant
    msgAry = cnvRange2Ary(shMsg, ROW_SH_MSG.IDX_DTL_BGN, rowMax, COL_SH_MSG.IDX_ID, COL_SH_MSG.IDX_MSG)
        
    '���b�Z�[�W�N���A
    Dim msg As String
    msg = ""

    '�z����T�[�`���Ăh�c����v����ꍇ�Ƀ��b�Z�[�W���擾
    Dim row As Long
    For row = LBound(msgAry) To UBound(msgAry)
        If Trim(msgAry(row, COL_SH_MSG.IDX_ID)) = id Then
            msg = msgAry(row, COL_SH_MSG.IDX_MSG)
            Exit For
        End If
    Next

    '���b�Z�[�W���擾�ł����ꍇ�A�u�������������u��
    If Not isBlank(msg) Then
        msg = getReplacedMessage(msg, replaceStrings)
    End If

    '���b�Z�[�W��Ԃ�
    getMsg = msg & "(" & id & ")"

    GoTo Fin

ErrorHandler:
    '�G���[����

Fin:
    '�㏈��
    If Not IsEmpty(msgAry) Then
        Erase msgAry
    End If

    '�G���[�����������ꍇ�͌ďo���ɓ�����
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If

End Function

'************************************************************************************
' �@�@�\    :���b�Z�[�W�̒u�������������u��
' ���@��    :in  msg                            ���b�Z�[�W
'           :in  replaceStrings                 �u������������z��
' �߁@�l    :���b�Z�[�W
' -----------------------------------------------------------------------------------
' ���@���@�@:2022/06/14  pecorimaru  Ver.1.0.00  �V�K�쐬
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
' �@�@�\    :�G���[�I�u�W�F�N�gRaise
' ���@��    :in  msg                            ���b�Z�[�W
'           :in  errSource                      �G���[�����ӏ�
'           :in  replaceStrings                 ���b�Z�[�W�u������������(�����w���)
' �߁@�l    :�G���[�I�u�W�F�N�g
'************************************************************************************
Public Sub raiseCommonErr( _
      ByVal msg As String _
    , ByVal errSource As String _
    , ParamArray replaceStrings() _
)

    '���O��
    '���b�Z�[�W �c�u���b�Z�[�WID:���b�Z�[�W���e�v�`��
    '���b�Z�[�W�h�c �c�uC_Mx0000�v�`��

    '���b�Z�[�W���AErr.Number���擾����(���b�Z�[�WID�̐��l���� + vbObjectError)
    Dim msgNo As Long
    msgNo = CLng(Mid(msg, 5, 4))
    msgNo = msgNo + vbObjectError

    '�G���[Raise
    Call Err.Raise(msgNo, errSource, ReplacePh(msg, replaceStrings))

End Sub

'************************************************************************************
' �@�@�\    :�A�v���P�[�V�����ݒ�(�����J�n)
' ���@��    :VBA���s���ɉ�ʕ`��⃆�[�U�[���삪�s���Ȃ��悤�ɐ���
' ���@��    :in      statusbarString            �X�e�[�^�X�o�[�\��������
'           :                                     ��"FALSE"�Ƃ���������͎w��s��
' �߁@�l    :�Ȃ�
'************************************************************************************
Public Sub setApplicationSettingsStart( _
    Optional ByVal statusbarString As String = vbNullString _
)

    If Not isProcessing Then
    
        '�����J�n�O�̐ݒ��ۑ�
        appSettingsSave = getAppSettings

        With Application
            
            '�ݒ�
            .EnableEvents = False                               '�C�x���g��}��
            .Interactive = False                                '���[�U������󂯕t���Ȃ�
            .ScreenUpdating = False                             '�`��X�V���~
            .Calculation = XlCalculation.xlCalculationManual    '�����Čv�Z���蓮�ɐݒ�
            .Cursor = XlMousePointer.xlWait                     '�J�[�\����WAIT�ɕύX

            '�X�e�[�^�X�o�[�\�������񂪎w�肳�ꂽ�ꍇ
            If Not StrPtr(statusbarString) = 0 Then
                '�X�e�[�^�X�o�[��ύX
                .DisplayStatusBar = True
                .StatusBar = statusbarString
            End If
        
        End With

        isProcessing = True
        isPausing = False
        
    End If

End Sub

'************************************************************************************
' �@�@�\    :�A�v���P�[�V�����ݒ�(�����I��)
' ���@��    :�ݒ�������J�n�O�ɖ߂�
' ���@��    :�Ȃ�
' �߁@�l    :�Ȃ�
'************************************************************************************
Public Sub setApplicationSettingsEnd()

    If isProcessing Then
    
        '�ݒ�������J�n�O�ɖ߂�
        '  �X�e�[�^�X�o�[�F�ύX����(�����J�n�O�̏�Ԃɖ߂�)
        '  �v�Z���@�F�ύX����(�����J�n�O�̏�Ԃɖ߂�)
        Call setAppSettings(appSettingsSave, True, True)

        isProcessing = False
        isPausing = False
    
    End If

End Sub

'************************************************************************************
' �@�@�\    :�A�v���P�[�V�����ݒ�(�ꎞ��~)
' ���@��    :in      calcModeChange             �v�Z���@�������J�n�O�̏�ԂɕύX����
' �߁@�l    :�Ȃ�
'************************************************************************************
Public Sub setApplicationSettingsPause( _
    Optional ByVal calcModeChange As Boolean = True _
)

    If (isProcessing) And (Not isPausing) Then
    
        '�ꎞ��~�O�̐ݒ��ۑ�
        appSettingsPauseSave = getAppSettings

        '�ݒ�������J�n�O�ɖ߂�
        '  �X�e�[�^�X�o�[�F�ύX���Ȃ�
        '  �v�Z���@�F�����J�n�O�̏�ԂɕύX���邩�ǂ����t���O�Ŕ��f
        Call setAppSettings(appSettingsSave, False, calcModeChange)

        '�ꎞ��~�t���O�n�m
        isPausing = True
    
    End If

End Sub

'************************************************************************************
' �@�@�\    :�A�v���P�[�V�����ݒ�(�Ďn��)
' ���@��    :�Ȃ�
' �߁@�l    :�Ȃ�
'************************************************************************************
Public Sub setApplicationSettingsReStart()

    If isPausing Then

         '�ݒ���ꎞ��~�O�ɖ߂�
         '  �X�e�[�^�X�o�[�F�ύX���Ȃ�
         '  �v�Z���@�F�ύX����(�ꎞ��~�O�̏�Ԃɖ߂�)
         Call setAppSettings(appSettingsPauseSave, False, True)
 
         '�ꎞ��~�t���O�n�e�e
         isPausing = False

    End If

End Sub

'************************************************************************************
' �@�@�\    :�A�v���P�[�V�����ݒ�(��ʍX�V)
' ���@��    :in      withAutoCalc               �Čv�Z�t���O
' �߁@�l    :�Ȃ�
'************************************************************************************
Public Sub setApplicationSettingsRefresh( _
    Optional ByVal withAutoCalc As Boolean = True _
)
        
    '���݂̐ݒ��ۑ�
    Dim tmpAppSettingsSave As AppSettingsType
    tmpAppSettingsSave = getAppSettings

    With Application
        
        '��ʍX�V
        .EnableEvents = True
        .Interactive = True
        .ScreenUpdating = True

        '�Čv�Z����
        If withAutoCalc Then
            .Calculate
        End If
    
    End With

    DoEvents

    '�ݒ��߂�
    '  �X�e�[�^�X�o�[�F�ύX���Ȃ�
    '  �v�Z���@�F�ύX���Ȃ�
    Call setAppSettings(tmpAppSettingsSave, False, False)

End Sub

'************************************************************************************
' �@�@�\    :���݂̃A�v���P�[�V�����ݒ���擾
' ���@��    :�Ȃ�
' �߁@�l    :���݂̃A�v���P�[�V�����ݒ�
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
' �@�@�\    :�A�v���P�[�V�����ݒ��ύX
' ���@��    :in      appSettingsValue           �A�v���P�[�V�����ݒ�l
'           :in      statusBarSet               �X�e�[�^�X�o�[��ݒ肷�邩
'           :in      calcModeSet                �v�Z���@��ݒ肷�邩
' �߁@�l    :�Ȃ�
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



