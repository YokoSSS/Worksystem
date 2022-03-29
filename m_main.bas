Attribute VB_Name = "m_main"
Option Explicit

'******************************************************************************
'FunctionName:mainFILENAMEGET
'Specifications：指定したフォルダ内のファイル名を取得しlistシート一覧へ出力
'                します。
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub mainFILENAMEGET()

  Dim Fdnfullpath As String
  
  On Error GoTo Err_Trap
  
  Call m_common.マクロ開始
  
  If CellDelete = False Then Exit Sub
  
  If filenameget(Fdnfullpath) = False Then Exit Sub
    Debug.Print "( Fdnfullpath: " & Fdnfullpath & ")"
  
  If extendget = False Then Exit Sub

  MsgBox "Processing Complete"
  
  Call m_common.マクロ終了
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  mainFILENAMEGET" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.マクロ終了
  
  End If

End Sub


'******************************************************************************
'FunctionName:mainFILERENAME
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub mainFILERENAME()

  Dim Fdnfullpath As String
  
  On Error GoTo Err_Trap
  
  Call m_common.マクロ開始
    
  If FileRename(Range("main_Fdnfullpath").Value) = False Then Exit Sub
    
  MsgBox "Done!!"

  Call m_common.マクロ終了
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  mainFILERENAME" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  Call m_common.マクロ終了
  
  End If

End Sub

'******************************************************************************
'FunctionName:mainFILEMOVE
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub mainFILEMOVE()

  Dim Fdnfullpath As String
  
  On Error GoTo Err_Trap
  
  Call m_common.マクロ開始
    
  If Filemove(Range("main_Fdnfullpath").Value) = False Then Exit Sub
    
  MsgBox "Done!!"

  Call m_common.マクロ終了
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  mainFILERENAME" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  Call m_common.マクロ終了
  
  End If

End Sub

