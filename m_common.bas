Attribute VB_Name = "m_common"
Option Explicit

'******************************************************************************
'FunctionName:MacroStart
'Specifications�FMacroStart���A�����X�s�[�h�����߂�ׂɖ��ʂȓ��������铮���
'��~������
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************

Sub MacroStart()
    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��
End Sub

'******************************************************************************
'FunctionName:MacroEnd
'Specifications�FMacroStart���A�����X�s�[�h�����߂�ׂɖ��ʂȓ��������铮���
'��~�����Ă������̂��ĉғ�������
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub Macroend()
    Application.StatusBar = False '�X�e�[�^�X�o�[������
    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n
End Sub


