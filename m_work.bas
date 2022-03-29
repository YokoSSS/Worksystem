Attribute VB_Name = "m_work"
Option Explicit

'******************************************************************************
'FunctionName:
'SpecificationsÅF
'ArgumentsÅFnothing
'ReturnValue:nothing
'NoteÅF
'******************************************************************************

Sub sample()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  ThisWorkbook.Save
  
  MsgBox "Done!!"

  Call m_common.MacroEnd
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  sample" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.MacroEnd
  
  End If

End Sub


