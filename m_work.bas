Attribute VB_Name = "m_work"
Option Explicit

'******************************************************************************
'FunctionName:CellDelete
'Specifications�Flist�ꗗ���������܂�
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function CellDelete() As Boolean
  
  Dim lastrow As Long
  Dim sr As Range
  
  On Error GoTo Err_Trap
  
  CellDelete = False

  With S_list

  lastrow = .Cells(Rows.Count, .Range("list_nid").Column).End(xlUp).Row

  If lastrow <> 1 Then

    Set sr = .Range("list_nid").Offset(1, 0)
    
    Debug.Print "(StartRange: " & sr.Address & ")"
    Debug.Print "(LastRange: " & _
      .Cells(lastrow, .Range("list_move���~").Column).Address & ")"

    .Range(sr, .Cells(lastrow, .Range("list_move���~").Column)).ClearContents
    Range("main_Fdnfullpath").ClearContents
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
    Debug.Print "FunctionName:CellDelete " & Err.Number & " " & Err.Description
    MsgBox "FunctionName:  CellDelete" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.�}�N���I��
  
  End If

End Function

'******************************************************************************
Function filenameget(Fdnfullpath As String) As Boolean

Dim FileCollection As Object
Dim FileList As Variant
Dim cnt As Long

  On Error GoTo Err_Trap

  filenameget = False

  MsgBox "�t�H���_���w�肵�ĉ������B"

  With Application.FileDialog(msoFileDialogFolderPicker)

    If .Show = True Then

      Fdnfullpath = .SelectedItems(1)

      Debug.Print "(getfoldername: " & Fdnfullpath & ")"

      Range("main_Fdnfullpath").Value = Fdnfullpath & "\"
    
    Else
      
      MsgBox "Finish processing"
      
      Call m_common.�}�N���I��
     
      Exit Function
    
    End If
  
  End With
  
  Set FileCollection = CreateObject("Scripting.FileSystemObject") _
    .GetFolder(Fdnfullpath).Files
  
  cnt = 0
  
  For Each FileList In FileCollection
    
    cnt = cnt + 1
    
    With S_list
      
      .Range("list_nid").Offset(cnt, 0) = cnt
      
      .Range("list_beforefilename").Offset(cnt, 0) = FileList.Name
    
    End With
  
  Next
  
  Set FileCollection = Nothing
  
  filenameget = True
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  filenameget" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.�}�N���I��
  
  End If

End Function

'******************************************************************************
'FunctionName:FileRename
'Specifications�F
'Arguments�FFdnfullpath/String
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function FileRename(ByRef Fdnfullpath As String) As Boolean
Dim lastrow As Long
Dim tarrg As Variant
Dim allr As Range
Dim strfilename As String

  On Error GoTo Err_Trap

  FileRename = False
  
  Debug.Print "(getfolderpath: " & Fdnfullpath & ")"
  
  With S_list
  
    lastrow = .Cells(Rows.Count, .Range("list_nid").Column).End(xlUp).Row - 1
   
    Debug.Print "(S_list lastrow: " & lastrow & ")"

    Set allr = .Range(.Range("list_beforefilename").Offset(1, 0), _
      .Range("list_beforefilename").Offset(lastrow, 0))
      Debug.Print "(allr: " & allr.Address & ")"
    
    End With
    
    For Each tarrg In allr

     Debug.Print "(file name before : " & tarrg & ")"
     Debug.Print "(file name after : " & tarrg.Offset(0, 1) & ")"
      
      If Dir(Fdnfullpath & tarrg) <> "" Then
        
        If tarrg <> "" And tarrg.Offset(0, 1) <> "" Then
          
          '�����̃t�@�C�������Ȃ��ꍇ���l�[������
          If Dir(Fdnfullpath & tarrg.Offset(0, 1)) = "" Then

            Name Fdnfullpath & tarrg As Fdnfullpath & tarrg.Offset(0, 1)

            tarrg.Offset(0, 2).Value = "Complete"

          '�����̃t�@�C����������ꍇ�x�������o�͂��A���l�[������
          ElseIf Dir(Fdnfullpath & tarrg.Offset(0, 1)) <> "" Then

            strfilename = InputBox("�����̃t�@�C���������݂��܂��B" & _
              "�t�@�C������ύX���Ă�������", "�t�@�C���l�[������", _
              tarrg.Offset(0, 1).Value)

            If strfilename <> "" Then

              Name Fdnfullpath & tarrg As Fdnfullpath & strfilename

              tarrg.Offset(0, 2).Value = "Complete"

            Else

              Name Fdnfullpath & tarrg As Fdnfullpath & tarrg.Offset(0, 1)

              tarrg.Offset(0, 2).Value = "Complete"

            End If

          End If
        
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
    MsgBox "FunctionName:  FileRename" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.�}�N���I��
  
  End If

End Function

'******************************************************************************
'FunctionName:extendget
'Specifications�F�g���q���擾���܂��B
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function extendget() As Boolean

Dim i As Long
Dim tarrg As String

  On Error GoTo Err_Trap

  extendget = False
  
  With S_list

    For i = 2 To .Cells(Rows.Count, .Range("list_beforefilename").Column).End(xlUp).Row

      tarrg = .Cells(i, .Range("list_beforefilename").Column).Value
      '�g���q������Ƃ��͏���������
      If InStrRev(tarrg, ".") > 0 Then
      
      .Cells(i, .Range("list_extend").Column).Value = _
        Mid(tarrg, InStrRev(tarrg, "."), _
        Len(tarrg) - InStrRev(tarrg, ".") + 1)
      
      End If

    Next i
  
  End With
  
  extendget = True
    
  Exit Function
   
Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  extendget" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
   
    Call m_common.�}�N���I��
  
  End If

End Function

'******************************************************************************
'FunctionName:FileMove
'Specifications�F
'Arguments�FFdnfullpath/String
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function Filemove(ByRef Fdnfullpath As String) As Boolean
Dim lastrow As Long
Dim tarrg As Variant
Dim allr As Range
Dim strfilename As String
Dim afterpath As String
Dim Comp As String
  On Error GoTo Err_Trap

  Filemove = False
  
  Debug.Print "(getfolderpath: " & Fdnfullpath & ")"
  
  With S_list
  
    lastrow = .Cells(Rows.Count, .Range("list_nid").Column).End(xlUp).Row - 1
   
    Debug.Print "(S_list lastrow: " & lastrow & ")"

    Set allr = .Range(.Range("list_beforefilename").Offset(1, 0), _
      .Range("list_beforefilename").Offset(lastrow, 0))
      Debug.Print "(allr: " & allr.Address & ")"
    
    For Each tarrg In allr

      Debug.Print "(file name before : " & tarrg & ")"
      Debug.Print "(file name after : " & tarrg.Offset(0, 1) & ")"
      
      If Dir(Fdnfullpath & tarrg) <> "" Then
        
        afterpath = .Cells(tarrg.Row, .Range("list_afterfilePath").Column)
        
        If tarrg <> "" And afterpath <> "" Then
          afterpath = afterpath & "\"
          '�����̃t�@�C�������Ȃ��ꍇ���l�[������
          If Dir(afterpath & tarrg) = "" Then

            Name Fdnfullpath & tarrg As afterpath & tarrg

            .Cells(tarrg.Row, .Range("list_move���~").Column).Value = "Complete"

          '�����̃t�@�C����������ꍇ�x�������o�͂��A���l�[������
          ElseIf Dir(Fdnfullpath & tarrg.Offset(0, 1)) <> "" Then

            strfilename = InputBox("�����̃t�@�C���������݂��܂��B" & _
              "�t�@�C������ύX���Ă�������", "�t�@�C���l�[������", _
              tarrg.Offset(0, 1).Value)

            If strfilename <> "" Then

              Name Fdnfullpath & tarrg As Fdnfullpath & strfilename

              .Cells(tarrg.Row, .Range("list_move���~").Column).Value = "Complete"

            Else

              Name Fdnfullpath & tarrg As Fdnfullpath & tarrg.Offset(0, 1)

              .Cells(tarrg.Row, .Range("list_move���~").Column).Value = "Complete"

            End If

          End If
        
        ElseIf tarrg = "" Then

            .Cells(tarrg.Row, .Range("list_move���~").Column).Value = _
            "Not Complete(Please enter before change file name.)"
        
        ElseIf .Cells(tarrg.Row, .Range("list_afterfilePath").Column).Value = "" Then

          .Cells(tarrg.Row, .Range("list_move���~").Column).Value = _
          "Not Complete(Please enter after change file name.)"
          
        End If

      Else
        
        .Cells(tarrg.Row, .Range("list_move���~").Column).Value = "Not Complete"
      
      End If
       
    Next
   
  End With
  
  Set allr = Nothing
   
  Filemove = True
   
    Exit Function
  
Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  FileRename" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.�}�N���I��
  
  End If

End Function


