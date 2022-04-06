Attribute VB_Name = "m_common"
Option Explicit

'******************************************************************************
'FunctionName:Macrostart
'Specifications�FStops unnecessary movements during Macrostart to increase
'processing speed.
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub Macrostart()
    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��
End Sub

'******************************************************************************
'FunctionName:Macroend
'Specifications�FRestarting operations that had been stopped to increase
'processing speed to increase wasteful movements during Macrostart.
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
    
    Call m_common.Macroend

    'Clear error
    Err.Clear
  
  End If

End Function

'******************************************************************************
'FunctionName:GET�t�@�C��
'Specifications�F�t�@�C���_�C�A���O�Ńt�@�C�����s�b�N�A�b�v
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function GET�t�@�C��(ByRef strFL As String) As Boolean

  On Error GoTo Err_Trap

  GET�t�@�C�� = False
  
  With Application.FileDialog(msoFileDialogFilePicker)
    
    If .Show = True Then
        
      strFL = .SelectedItems(1)
       
      Debug.Print "GET�t�@�C����:�@" & strFL

    Else
      
      MsgBox "�L�����Z�����܂����B�������I�����܂��B", vbInformation, "�����I��"
    
    End If

  End With

  GET�t�@�C�� = True

  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear

    Call m_common.Macroend
  
  End If

End Function

'******************************************************************************
'FunctionName:�w�肵���p�X�Ńt�H���_����
'Specifications�F�w�肵���p�X�Ńt�H���_����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function �w�肵���p�X�Ńt�H���_����(ByRef sFdPath As String) As Boolean
    
  Dim FSO As Object
  
  On Error GoTo Err_Trap
  
  �w�肵���p�X�Ńt�H���_���� = False
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FSO.CreateFolder sFdPath
  
  Set FSO = Nothing

  �w�肵���p�X�Ńt�H���_���� = True
    
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear

    Call m_common.Macroend
  
  End If

End Function

'******************************************************************************
'FunctionName:�t�H���_�̑��ݗL��
'Specifications�F�t�H���_�����݂��邩�ǂ����𒲂ׂ�Function�v���V�[�W��
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Function FolderExists(folder_path As String) As Boolean
  
  On Error GoTo Err_Trap
  
  If Dir(folder_path, vbDirectory) = "" Then
    
    FolderExists = False
  
  Else
    
    FolderExists = True
  
  End If
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear

    Call m_common.Macroend
  
  End If

End Function
