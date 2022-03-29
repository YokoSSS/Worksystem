Attribute VB_Name = "m_work"
Option Explicit
'******************************************************************************
'FunctionName:
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************

'******************************************************************************
'関数名：MakeDictionary
'仕様：Work2シートの入力情報によりDictionaryListを作成する
'引数：dic : Dictionary
'返り値：string　download stritems:string / changefilename:string
'NOTE：なし
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
  
  'オブジェクトの生成
  Set ws = ThisWorkbook.Worksheets("work2")

  '最終行取得
  lastrow = ws.Cells(Rows.Count, ws.Range("WORK2_IDN").Column).End(xlUp).Row

  'work2シートのヘッダー行をのぞくので-1する
  For i = 1 To lastrow - 1

    '(Key)
    strKeys = ws.Range("WORK2_IDN").Offset(i, 0).Value & "_" & _
      ws.Range("WORK2_Keys").Offset(i, 0).Value
      
    '(item)
    stritems = ws.Range("WORK2_Items").Offset(i, 0).Value
      
    'まだ登録をされていない場合登録をする
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
    
    MsgBox "Function MakeDictionary：" & Err.Number & " " & Err.Description & vbCrLf & _
    "処理を終了します。", vbOKOnly, "MakeDictionary: Failure"

    'Clear error
    Err.Clear
  
    Call マクロ終了
  
  End If

End Function

'******************************************************************************
'FunctionName:OutputDictionary
'Specifications：生成したDictionryオブジェクトよりFileListシートへ出力する
'Arguments：dictionary : dic
'ReturnValue:OutputDictionary:Boolean true / false
'Note：
'******************************************************************************
Function OutputDictionary(ByVal Dic As Dictionary) As Boolean

  On Error GoTo Err_Trap
  
  Call m_common.マクロ開始

  OutputDictionary = False
  
  Dim wsLS As Worksheet
  Dim i As Long, lastrow As Long
  
    'FileListシートのオブジェクト生成
  Set wsLS = ThisWorkbook.Worksheets("FileList")
  
  '最終行取得
  lastrow = wsLS.Cells(Rows.Count, wsLS.Range("FileList_IDN").Column).End(xlUp).Row
  Debug.Print "function: OutputDictionary 最終行取得:(" & lastrow & ")"
  Debug.Print "function: OutputDictionary dicのitem数:(" & Dic.Count & ")"
  
  'FileListシートへ書き出す
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
    
    MsgBox "Function OutputDictionary：" & Err.Number & " " & _
    Err.Description & vbCrLf & _
    "処理を終了します。", vbOKOnly, "OutputDictionary: Failure"

    'Clear error
    Err.Clear
  
    Call マクロ終了
  
  End If

End Function


