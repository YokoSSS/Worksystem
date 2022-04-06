Attribute VB_Name = "m_work"
Option Explicit

'******************************************************************************
'FunctionName:CellDelete
'Specifications：CellDelete
'Arguments：nothing
'ReturnValue:boolean true / false
'Note：Nothing
'******************************************************************************
Function CellDelete() As Boolean
  
  On Error GoTo Err_Trap
  
  CellDelete = False
  
  Dim lastrow As Long
  Dim sr As Range
    
  With S_rename
    
    lastrow = .Cells(Rows.Count, S_rename.Range("NID").Column).End(xlUp).Row
    
    If lastrow <> 2 Then
      
      Set sr = .Range("NID").Offset(1, 0)
      
      Debug.Print "(StartRange: " & sr.Address & ")"
      Debug.Print "(LastRange: " & _
      .Cells(lastrow, .Range("RenameComplete").Column).Address & ")"
      
      .Range(sr, .Cells(lastrow, .Range("extend").Column)).ClearContents
     
    Else
       
      CellDelete = True
      
      Exit Function
     
    End If
    
  End With
    
  Set sr = Nothing
  
  CellDelete = True
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  CellDelete" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Function

'******************************************************************************
'FunctionName:filenameget
'Specifications：Outputs the file names in the selected folder to the rename sheet.
'Arguments：Fdnfullpath / string (selected folder name)
'ReturnValue: boolean true / false
'Note：
'******************************************************************************
Function filenameget(ByVal Fdnfullpath As String) As Boolean
  
  On Error GoTo Err_Trap
  
  filenameget = False
  
  Dim FileCollection As Object
  Dim FileList As Variant
  Dim cnt As Long


  Fdnfullpath = Range("Fdnfullpath").Value & "\"
  
  Set FileCollection = CreateObject("Scripting.FileSystemObject") _
          .GetFolder(Fdnfullpath).Files
  
  cnt = 0
  For Each FileList In FileCollection
      cnt = cnt + 1
      With S_rename
          .Range("NID").Offset(cnt, 0) = cnt
          .Range("CurrentFilename").Offset(cnt, 0) = FileList.Name
      End With
  Next
  
  Set FileCollection = Nothing
  
  filenameget = True
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "FunctionName:  filenameget" & vbCrLf & _
      Err.Number & " " & Err.Description
    
    MsgBox "FunctionName:  filenameget" & vbCrLf & Err.Number & " " & _
      Err.Description, vbOKOnly, "filenameget : Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Function

'******************************************************************************
'FunctionName:extendget
'Specifications：
'Arguments：nothing
'ReturnValue: Boolean true / false
'Note：
'******************************************************************************
Function extendget() As Boolean
  
  On Error GoTo Err_Trap
  
  extendget = False
  
  Dim i As Long
  Dim tarrg As String
  
  With S_rename
    
    For i = 3 To .Cells(Rows.Count, 2).End(xlUp).Row
      
      tarrg = .Cells(i, 2).Value
      
      .Cells(i, 5).Value = Mid(tarrg, InStrRev(tarrg, "."), _
        Len(tarrg) - InStrRev(tarrg, ".") + 1)
    
    Next i
  
  End With
  
  extendget = True

  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  sample" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Function

'******************************************************************************
'FunctionName:FileRename
'Specifications：FileRename
'Arguments：Fdnfullpath / string (selected folder name)
'ReturnValue: boolean true / false
'Note：
'******************************************************************************
Function FileRename(ByVal Fdnfullpath As String) As Boolean
  
  On Error GoTo Err_Trap
  
  FileRename = False
  
  Dim lastrow As Long
  Dim tarrg As Variant
  Dim allr As Range

  Debug.Print "(getfolderpath: " & Fdnfullpath & ")"
  
  S_rename.Activate

  lastrow = S_rename.Cells(Rows.Count, S_rename.Range("NID").Column).End(xlUp).Row - 2
  
  Debug.Print "(S_rename lastrow: " & lastrow & ")"

  Set allr = S_rename.Range(S_rename.Range("CurrentFilename").Offset(1, 0), _
    S_rename.Range("CurrentFilename").Offset(lastrow, 0))
    Debug.Print "(allr: " & allr.Address & ")"

  For Each tarrg In allr
      
    Debug.Print "(tarrg: " & tarrg & ")"
    Debug.Print "(tarrg: " & tarrg.Offset(0, 1) & ")"
    Debug.Print "(tarrg: " & tarrg.Address & ")"
    
    If Dir(Fdnfullpath & tarrg) <> "" Then
        
      If tarrg <> "" And tarrg.Offset(0, 1) <> "" Then
        
        Name Fdnfullpath & tarrg As Fdnfullpath & tarrg.Offset(0, 1)
        
        tarrg.Offset(0, 2).Value = "Complete"
      
      ElseIf tarrg = "" Then
        
        tarrg.Offset(0, 2).Value = "Not Complete(Please enter before change file name.)"
      
      ElseIf tarrg.Offset(0, 1) = "" Then
        
        tarrg.Offset(0, 2).Value = "Not Complete(Please enter after change file name.)"
      
      End If
    
    Else
      
      tarrg.Offset(0, 2).Value = "Not Complete"
    
    End If
      
  Next
  
  Set allr = Nothing
  
  FileRename = True
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    
    MsgBox "FunctionName:  FileRename" & vbCrLf & Err.Number & " " & _
      Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Function


