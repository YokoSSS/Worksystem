Attribute VB_Name = "m_work"
Option Explicit

'******************************************************************************
'FunctionName:GetDataFromADODBRS
'Specifications�F���R�[�h�Z�b�g����CopyFromRecordset���\�b�h�Ńf�[�^�擾
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub GetDataFromADODBRS()

  Dim MySql As String, MyPath As String
  Dim i  As Integer
  Dim Conn As ADODB.Connection
  Dim Rst As ADODB.Recordset
  
  On Error GoTo Err_Trap
  
  Call m_common.�}�N���J�n

  Set Conn = New ADODB.Connection
  
  MyPath = "c:\�̔��Ǘ�.mdb" '�f�[�^�x�[�X�̎w��
  Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                                  & "Data Source=" & MyPath & ";"
  Conn.Open '�ڑ��̊m��
  
  '�̔��Ǘ�.mdb���̔���e�[�u����肷�ׂẴf�[�^���擾�B
  '���t�t�B�[���h����ɍ~���\�[�g
  MySql = "select * from ���� order by ���t desc;"
  
  Set Rst = New ADODB.Recordset
  Rst.Open MySql, Conn, adOpenStatic, adLockReadOnly, adCmdText
  
  '�t�B�[���h���̏����o��
  For i = 0 To Rst.Fields.Count - 1
  ActiveSheet.Cells(1, i + 1).Value = Rst.Fields(i).Name
  Next i
  'CopyFromRecordset���\�b�h�Ŋ�Z�����w�肵�ăf�[�^�̏����o��
  ActiveSheet.Range("a2").CopyFromRecordset Rst
  
  Rst.Close: Conn.Close
  Set Rst = Nothing: Set Conn = Nothing
  
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
'******************************************************************************
'FunctionName:GetDataByQueryTable
'Specifications�F�f�[�^�x�[�X�N�G����ADODB���R�[�h�Z�b�g���g�p����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub GetDataByQueryTable()

  Dim QT As QueryTable
  Dim MySql As String, MyPath As String
  Dim Conn As ADODB.Connection
  Dim Rst As ADODB.Recordset
  
  On Error GoTo Err_Trap
  
  Call m_common.�}�N���J�n

  Set Conn = New ADODB.Connection
  
  MyPath = "c:\�̔��Ǘ�.mdb"
  Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                                  & "Data Source=" & MyPath & ";"
  Conn.Open
  
  '�̔��Ǘ�.mdb���̔���e�[�u�����t�B�[���h���w�肵��
  '�f�[�^���擾�B�ڋqID�t�B�[���h����ɏ����\�[�g
  MySql = "select �ڋqID,���iID,��,�P��" _
         & " from ���� order by �ڋqID ASC;"
  
  Set Rst = New ADODB.Recordset
  Rst.Open MySql, Conn, adOpenStatic, adLockReadOnly, adCmdText
  
  Set QT = ActiveSheet.QueryTables.Add _
      (Connection:=Rst, Destination:=Range("A1"))
  QT.Name = "MyQuery"
  QT.Refresh
  
  Rst.Close: Conn.Close
  Set Rst = Nothing: Set Conn = Nothing

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

