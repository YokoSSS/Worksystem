Attribute VB_Name = "m_work"
Option Explicit

'******************************************************************************
'FunctionName:GetDataFromADODBRS
'Specifications：レコードセットからCopyFromRecordsetメソッドでデータ取得
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub GetDataFromADODBRS()

  Dim MySql As String, MyPath As String
  Dim i  As Integer
  Dim Conn As ADODB.Connection
  Dim Rst As ADODB.Recordset
  
  On Error GoTo Err_Trap
  
  Call m_common.マクロ開始

  Set Conn = New ADODB.Connection
  
  MyPath = "c:\販売管理.mdb" 'データベースの指定
  Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                                  & "Data Source=" & MyPath & ";"
  Conn.Open '接続の確立
  
  '販売管理.mdb内の売上テーブルよりすべてのデータを取得。
  '日付フィールドを基準に降順ソート
  MySql = "select * from 売上 order by 日付 desc;"
  
  Set Rst = New ADODB.Recordset
  Rst.Open MySql, Conn, adOpenStatic, adLockReadOnly, adCmdText
  
  'フィールド名の書き出し
  For i = 0 To Rst.Fields.Count - 1
  ActiveSheet.Cells(1, i + 1).Value = Rst.Fields(i).Name
  Next i
  'CopyFromRecordsetメソッドで基準セルを指定してデータの書き出し
  ActiveSheet.Range("a2").CopyFromRecordset Rst
  
  Rst.Close: Conn.Close
  Set Rst = Nothing: Set Conn = Nothing
  
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
'******************************************************************************
'FunctionName:GetDataByQueryTable
'Specifications：データベースクエリでADODBレコードセットを使用する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub GetDataByQueryTable()

  Dim QT As QueryTable
  Dim MySql As String, MyPath As String
  Dim Conn As ADODB.Connection
  Dim Rst As ADODB.Recordset
  
  On Error GoTo Err_Trap
  
  Call m_common.マクロ開始

  Set Conn = New ADODB.Connection
  
  MyPath = "c:\販売管理.mdb"
  Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                                  & "Data Source=" & MyPath & ";"
  Conn.Open
  
  '販売管理.mdb内の売上テーブルよりフィールドを指定して
  'データを取得。顧客IDフィールドを基準に昇順ソート
  MySql = "select 顧客ID,商品ID,個数,単価" _
         & " from 売上 order by 顧客ID ASC;"
  
  Set Rst = New ADODB.Recordset
  Rst.Open MySql, Conn, adOpenStatic, adLockReadOnly, adCmdText
  
  Set QT = ActiveSheet.QueryTables.Add _
      (Connection:=Rst, Destination:=Range("A1"))
  QT.Name = "MyQuery"
  QT.Refresh
  
  Rst.Close: Conn.Close
  Set Rst = Nothing: Set Conn = Nothing

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

