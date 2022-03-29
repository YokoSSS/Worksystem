Attribute VB_Name = "m_main"
Option Explicit

'******************************************************************************
'FunctionName:
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************

Sub sample()
  
  On Error GoTo Err_Trap
  
  Call m_common.マクロ開始

  MsgBox "Done!!"

  Call m_common.マクロ終了
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  sample" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.マクロ終了
  
  End If

End Sub

