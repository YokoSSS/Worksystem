Attribute VB_Name = "m_common"
Option Explicit

'******************************************************************************
'FunctionName:�}�N���J�n
'Specifications�F�}�N���J�n���A�����X�s�[�h�����߂�ׂɖ��ʂȓ��������铮���
'��~������
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************

Sub �}�N���J�n()
    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��
End Sub

'******************************************************************************
'FunctionName:�}�N���I��
'Specifications�F�}�N���J�n���A�����X�s�[�h�����߂�ׂɖ��ʂȓ��������铮���
'��~�����Ă������̂��ĉғ�������
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub �}�N���I��()
    Application.StatusBar = False '�X�e�[�^�X�o�[������
    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n
End Sub


