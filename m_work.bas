Attribute VB_Name = "m_work"
Option Explicit
'******************************************************************************
'FunctionName:
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************

'******************************************************************************
'�֐����FMakeDictionary
'�d�l�FWork2�V�[�g�̓��͏��ɂ��DictionaryList���쐬����
'�����Fdic : Dictionary
'�Ԃ�l�Fstring�@download stritems:string / changefilename:string
'NOTE�F�Ȃ�
'******************************************************************************
Function MakeDictionary(ByRef Dic As Dictionary) As Boolean

  On Error GoTo Err_Trap
  
  MakeDictionary = False
  
  Dim ws As Worksheet
  Dim i As Long, lastrow As Long
  Dim strKeys As String, stritems As String
 
  strKeys = ""
  stritems = ""
  Debug.Print "function: MakeDictionary strKeys:(" & strKeys & ")"
  Debug.Print "function: MakeDictionary stritems:(" & stritems & ")"
  
  '�I�u�W�F�N�g�̐���
  Set ws = ThisWorkbook.Worksheets("work2")

  '�ŏI�s�擾
  lastrow = ws.Cells(Rows.Count, ws.Range("WORK2_IDN").Column).End(xlUp).Row

  'work2�V�[�g�̃w�b�_�[�s���̂����̂�-1����
  For i = 1 To lastrow - 1

    '(Key)
    strKeys = ws.Range("WORK2_IDN").Offset(i, 0).Value & "_" & _
      ws.Range("WORK2_Keys").Offset(i, 0).Value
      
    '(item)
    stritems = ws.Range("WORK2_Items").Offset(i, 0).Value
      
    '�܂��o�^������Ă��Ȃ��ꍇ�o�^������
    If Dic.Exists(strKeys) = False Then
        
        Dic.Add strKeys, stritems

    Else

        Dic.Item(stritems) = Dic.Item(stritems)
    
    End If

  Next i
  
  Set ws = Nothing
  
  MakeDictionary = True
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print " MakeDictionary: Error" & Err.Number & " " & Err.Description
    
    MsgBox "Function MakeDictionary�F" & Err.Number & " " & Err.Description & vbCrLf & _
    "�������I�����܂��B", vbOKOnly, "MakeDictionary: Failure"

    'Clear error
    Err.Clear
  
    Call �}�N���I��
  
  End If

End Function

'******************************************************************************
'FunctionName:OutputDictionary
'Specifications�F��������Dictionry�I�u�W�F�N�g���FileList�V�[�g�֏o�͂���
'Arguments�Fdictionary : dic
'ReturnValue:OutputDictionary:Boolean true / false
'Note�F
'******************************************************************************
Function OutputDictionary(ByVal Dic As Dictionary) As Boolean

  On Error GoTo Err_Trap
  
  Call m_common.�}�N���J�n

  OutputDictionary = False
  
  Dim wsLS As Worksheet
  Dim i As Long, lastrow As Long
  
    'FileList�V�[�g�̃I�u�W�F�N�g����
  Set wsLS = ThisWorkbook.Worksheets("FileList")
  
  '�ŏI�s�擾
  lastrow = wsLS.Cells(Rows.Count, wsLS.Range("FileList_IDN").Column).End(xlUp).Row
  Debug.Print "function: OutputDictionary �ŏI�s�擾:(" & lastrow & ")"
  Debug.Print "function: OutputDictionary dic��item��:(" & Dic.Count & ")"
  
  'FileList�V�[�g�֏����o��
  For i = 0 To Dic.Count - 1
  
    wsLS.Range("FileList_IDN").Offset(i + lastrow, 0).Value = _
      Mid(Dic.Keys(i), 1, InStr(Dic.Keys(i), "_") - 1)
    
    wsLS.Range("FileList_dlpic").Offset(i + lastrow, 0).Value = _
      Mid(Dic.Keys(i), InStr(Dic.Keys(i), "_") + 1, _
      Len(Dic.Keys(i)) - InStr(Dic.Keys(i), "_"))
    
    wsLS.Range("FileList_chFl").Offset(i + lastrow, 0).Value = _
      Dic.items(i)
  
  Next i
  
  OutputDictionary = True
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print " OutputDictionary: Error" & Err.Number & " " & Err.Description
    
    MsgBox "Function OutputDictionary�F" & Err.Number & " " & _
    Err.Description & vbCrLf & _
    "�������I�����܂��B", vbOKOnly, "OutputDictionary: Failure"

    'Clear error
    Err.Clear
  
    Call �}�N���I��
  
  End If

End Function


