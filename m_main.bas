Attribute VB_Name = "m_main"
Option Explicit

'******************************************************************************
'FunctionName:mainFILENAMEGET
'Specifications�F�w�肵���t�H���_���̃t�@�C�������擾��list�V�[�g�ꗗ�֏o��
'                ���܂��B
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub mainFILENAMEGET()

  Dim Fdnfullpath As String
  
  On Error GoTo Err_Trap
  
  Call m_common.�}�N���J�n
  
  If CellDelete = False Then Exit Sub
  
  If filenameget(Fdnfullpath) = False Then Exit Sub
    Debug.Print "( Fdnfullpath: " & Fdnfullpath & ")"
  
  If extendget = False Then Exit Sub

  MsgBox "Processing Complete"
  
  Call m_common.�}�N���I��
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  mainFILENAMEGET" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.�}�N���I��
  
  End If

End Sub


'******************************************************************************
'FunctionName:mainFILERENAME
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub mainFILERENAME()

  Dim Fdnfullpath As String
  
  On Error GoTo Err_Trap
  
  Call m_common.�}�N���J�n
    
  If FileRename(Range("main_Fdnfullpath").Value) = False Then Exit Sub
    
  MsgBox "Done!!"

  Call m_common.�}�N���I��
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  mainFILERENAME" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  Call m_common.�}�N���I��
  
  End If

End Sub

'******************************************************************************
'FunctionName:mainFILEMOVE
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub mainFILEMOVE()

  Dim Fdnfullpath As String
  
  On Error GoTo Err_Trap
  
  Call m_common.�}�N���J�n
    
  If Filemove(Range("main_Fdnfullpath").Value) = False Then Exit Sub
    
  MsgBox "Done!!"

  Call m_common.�}�N���I��
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  mainFILERENAME" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  Call m_common.�}�N���I��
  
  End If

End Sub

