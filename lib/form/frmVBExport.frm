VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVBExport 
   Caption         =   "VBAモジュールのエクスポート"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   105
   ClientWidth     =   8565.001
   OleObjectBlob   =   "frmVBExport.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmVBExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************************
' 機　能    :VBAモジュールのエクスポート
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/10/18  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************

'************************************************************************************
' 機　能    :ユーザーフォーム初期処理
' 説　明    :開いているブック名を全て取得し、[ワークブック選択]にセット
' 引　数    :なし
' 戻　値    :なし
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/10/18  J-Tam  Ver.1.0.00  新規作成
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
' 機　能    :参照ボタン（ファイル参照）押下時処理
' 説　明    :ファイル選択ダイアログを開き、選択したファイルをワークブックにセット
' 引　数    :なし
' 戻　値    :なし
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/10/18  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************
Private Sub btnRefBkExpt_Click()
    With frmVBExport.txtBkExpt
        .Value = showFilePickerDialog("ブックを選択", .Value, "*.xlsm;*.xlam")
    End With
End Sub

'************************************************************************************
' 機　能    :参照ボタン（保存先フォルダ）押下時処理
' 説　明    :フォルダ選択ダイアログを開き、選択したフォルダを保存先フォルダにセット
' 引　数    :なし
' 戻　値    :なし
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/10/18  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************
Private Sub btnRefSavePath_Click()
    With frmVBExport.txtSavePath
        .Value = showFolderPickerDialog("保存先を選択", .Value)
    End With
End Sub

'************************************************************************************
' 機　能    :オプションボタン（現在開いているブック）押下時処理
' 引　数    :なし
' 戻　値    :なし
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/10/18  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************
Private Sub optOpening_Click()
    frmVBExport.cmbBkExpt.Enabled = True
    frmVBExport.txtBkExpt.Enabled = False
    frmVBExport.btnRefBkExpt.Enabled = False
End Sub

'************************************************************************************
' 機　能    :オプションボタン（現在開いているブック）押下時処理
' 引　数    :なし
' 戻　値    :なし
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/10/18  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************
Private Sub optRefFile_Click()
    frmVBExport.cmbBkExpt.Enabled = False
    frmVBExport.txtBkExpt.Enabled = True
    frmVBExport.btnRefBkExpt.Enabled = True
End Sub

'************************************************************************************
' 機　能    :エクスポートボタン押下時処理
' 引　数    :なし
' 戻　値    :なし
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/10/18  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************
Private Sub btnExecExpt_Click()

    Dim module As VBComponent
    Dim moduleList As VBComponents
    Dim bkExpt As Workbook

    On Error GoTo ErrorHandler

    '************************************************************
    ' 初期処理
    '************************************************************

    '処理実行フラグを初期化
    Dim isJobed As Boolean
    isJobed = False

    '入力チェック
    If Not isValidInput() Then
        GoTo Fin
    End If

    'エクスポート対象ブックを取得
    If frmVBExport.optOpening = True Then
        Set bkExpt = Workbooks(frmVBExport.cmbBkExpt.Value)
    Else
        Set bkExpt = Workbooks.Open(frmVBExport.txtBkExpt.Value)
    End If
    
    'ブックのモジュール一覧を取得
    Set moduleList = bkExpt.VBProject.VBComponents
    
    '************************************************************
    ' メイン処理
    '************************************************************
    
    'VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList

        Dim isExpt As Boolean          'エクスポート対象フラグ
        Dim ext As String              '拡張子
        
        'モジュールごとに対象判定と拡張子を分岐
        Select Case module.Type
            
            'クラスモジュール
            Case vbext_ct_ClassModule

                isExpt = frmVBExport.chkClsMod.Value
                ext = "cls"

            'ユーザーフォーム
            Case vbext_ct_MSForm
                
                isExpt = frmVBExport.chkUserForm.Value
                ext = "frm"

            '標準モジュール
            Case vbext_ct_StdModule

                isExpt = frmVBExport.chkStdMod.Value
                ext = "bas"
                
            '上記以外
            Case Else
            
                isExpt = frmVBExport.chkBkSh.Value
                ext = "cls"
        
        End Select

        '対象モジュールにチェックがある場合、エクスポート
        If isExpt Then
            Dim exptPath As String
            exptPath = frmVBExport.txtSavePath.Value & "\" & module.Name & "." & ext
        
            '上書きチェック
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

    'ブックをクローズ
    If frmVBExport.optOpening = False Then
        bkExpt.Close
    End If

    '正常終了
    If isJobed Then
        Call showMsgBox(getMsg("MI0002", "エクスポート"), xlInfo)
    Else
        Call showMsgBox(getMsg("MI0003", "エクスポート"), xlInfo)
    End If
    GoTo Fin

ErrorHandler:
    '************************************************************
    ' エラー処理
    '************************************************************

    '異常終了
    Call showSystemErrorMsg(Err, ThisWorkbook.Name)
    Resume Fin

Fin:
    '************************************************************
    ' 後処理
    '************************************************************

    Set bkExpt = Nothing
    Set moduleList = Nothing
    Set module = Nothing

End Sub

'************************************************************************************
' 機　能    :入力チェック
' 引　数    :なし
' 戻　値    :なし
' -----------------------------------------------------------------------------------
' 履　歴　　:2022/10/18  J-Tam  Ver.1.0.00  新規作成
'************************************************************************************
Private Function isValidInput() As Boolean

    '戻り値をクリア
    isValidInput = False

    On Error GoTo ErrorHandler

    '************************************************************
    ' メイン処理
    '************************************************************

    With frmVBExport

    '----ワークブック選択----

        If frmVBExport.optOpening = True Then

            If isBlank(.cmbBkExpt.Value) Then
                Call showMsgBox(getMsg("ME0002", "ワークブック"), xlWarning)
                isValidInput = False
                GoTo Fin
            End If
    
        Else
    
            If isBlank(.txtBkExpt.Value) Then
                Call showMsgBox(getMsg("ME0001", "ワークブック"), xlWarning)
                isValidInput = False
                GoTo Fin
            End If
    
        End If
    
    '----保存先フォルダ----
    
        If isBlank(.txtSavePath.Value) Then
            Call showMsgBox(getMsg("ME0001", "保存先フォルダ"), xlWarning)
            isValidInput = False
            GoTo Fin
        End If
    
        If Not existsFolder(.txtSavePath.Value) Then
            Call showMsgBox(getMsg("ME0003", "保存先フォルダ", .txtSavePath.Value), xlWarning)
            isValidInput = False
            GoTo Fin
        End If

    End With

    '正常終了
    isValidInput = True
    GoTo Fin

ErrorHandler:
    '************************************************************
    ' エラー処理
    '************************************************************

    '異常終了
    Call showSystemErrorMsg(Err, ThisWorkbook.Name)
    isValidInput = False
    Resume Fin

Fin:
    '************************************************************
    ' 後処理
    '************************************************************

End Function
