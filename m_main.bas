Attribute VB_Name = "m_main"
Option Explicit

'******************************************************************************
'FunctionName:
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************

Sub sample()
  
  On Error GoTo Err_Trap
  
  Call m_common.�}�N���J�n

  MsgBox "Done!!"

  Call m_common.�}�N���I��
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  sample" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.�}�N���I��
  
  End If

End Sub

