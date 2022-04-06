Attribute VB_Name = "m_main"
Option Explicit

'******************************************************************************
'FunctionName:mainwork001
'Specifications：Outputs the file names in the selected folder to the rename
'sheet.
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************

Sub mainwork001()
  
  On Error GoTo Err_Trap
  
  Call m_common.Macrostart

  If CellDelete = False Then
  
    MsgBox "An error occurred while deleting data on the RENAME sheet." & vbCrLf & _
    "The process is terminated.", vbCritical, "CellDelete Error"
    
    Call m_common.Macroend
    
    Exit Sub
    
  End If

  If filenameget(Range("Fdnfullpath")) = False Then
  
    MsgBox "An error occurred while outputting the file name." & vbCrLf & _
    "The process is terminated.", vbCritical, "filenameget Error"
    
    Call m_common.Macroend
    
    Exit Sub
    
  End If
        
  Debug.Print "( Fdnfullpath: " & Range("Fdnfullpath") & ")"
    
  Call extendget

  ThisWorkbook.Save
  
  MsgBox "Processing Complete"

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  mainwork001" & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "mainwork001 Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:mainwork002
'Specifications：Convert file names
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub mainwork002()
  
  On Error GoTo Err_Trap
  
  Call m_common.Macrostart

  If FileRename(Range("Fdnfullpath").Value & "\") = False Then
  
    MsgBox "An error occurred while renaming a file in a folder." & vbCrLf & _
    "The process is terminated.", vbCritical, "CellDelete Error"
    
    Call m_common.Macroend
    
    Exit Sub
    
  End If
  
  
  ThisWorkbook.Save
  
  MsgBox "Processing Complete"

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "FunctionName:  mainwork002" & vbCrLf & _
    Err.Number & " " & Err.Description
    
    MsgBox "FunctionName:  mainwork002" & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "mainwork002 Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************

