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
Sub MacroEnd()
    Application.StatusBar = False '�X�e�[�^�X�o�[������
    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n
End Sub

'******************************************************************************
'FunctionName:GET�t�H���_
'Specifications�F�t�H���_�_�C�A���O�Ńt�H���_���s�b�N�A�b�v
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
'Function GET�t�H���_(ByRef strFD As String) As Boolean
Function GET�t�H���_()
Dim strFD As String
  On Error GoTo Err_Trap
  
  GET�t�H���_ = False
  
  With Application.FileDialog(msoFileDialogFolderPicker)
      
    If .Show = True Then
        
      strFD = .SelectedItems(1)
    
    Else

      MsgBox "�L�����Z�����܂����B�������I�����܂��B", vbInformation, "�����I��"

      Exit Function

    End If
  
  End With

  GET�t�H���_ = True

  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"
    
    Call m_common.MacroEnd

    'Clear error
    Err.Clear
  
  End If

End Function

'******************************************************************************
'FunctionName:GETFile
'Specifications�FPick a file in the file dialog up.
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function GETFile(ByRef strFL As String) As Boolean

  On Error GoTo Err_Trap

  GETFile = False
  
  With Application.FileDialog(msoFileDialogFilePicker)
    
    If .Show = True Then
        
      strFL = .SelectedItems(1)
       
      Debug.Print "GETFileName :�@" & strFL

    Else
      
      MsgBox "Canceled. The process is terminated.", vbInformation, _
      "End of processing"
    
    End If

  End With

  GETFile = True

  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "GETFile: " & Err.Number & " " & Err.Description
    MsgBox "GETFile: " & Err.Number & " " & Err.Description, vbOKOnly, _
    "GETFile: Error"

    'Clear error
    Err.Clear

    Call m_common.MacroEnd
  
  End If

End Function

'******************************************************************************
'FunctionName:CreateFolder
'Specifications�FCreate a folder with the specified path
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function CreateFolder(ByRef sFdPath As String) As Boolean
    
  Dim FSO As Object
  
  On Error GoTo Err_Trap
  
  CreateFolder = False
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FSO.CreateFolder sFdPath
  
  Set FSO = Nothing

  CreateFolder = True
    
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear

    Call m_common.MacroEnd
  
  End If

End Function

'******************************************************************************
'FunctionName:ExistenceOrNonexistenceFolders
'Specifications�FExistence or non-existence of folders
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function ExistenceOrNonexistenceFolders(folder_path As String) As Boolean
  
  On Error GoTo Err_Trap
  
  If Dir(folder_path, vbDirectory) = "" Then
    
    ExistenceOrNonexistenceFolders = False
  
  Else
    
    ExistenceOrNonexistenceFolders = True
  
  End If
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "ExistenceOrNonexistenceFolders: " & Err.Number & " " & _
    Err.Description
    MsgBox "ExistenceOrNonexistenceFolders : " & Err.Number & " " & _
    Err.Description, vbOKOnly, "ExistenceOrNonExistenceFolders: Error"

    'Clear error
    Err.Clear

    Call m_common.MacroEnd
  
  End If

End Function
