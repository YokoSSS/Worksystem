Attribute VB_Name = "m_main"
Option Explicit

'******************************************************************************
'FunctionName:sampleDictionary
'Specifications：Dictionaryオブジェクトに格納されたキーと要素を、格納された順に
'出力していきます。
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub sampleDictionary()
  
  On Error GoTo Err_Trap
  
  Call m_common.マクロ開始
  
  Dim strKeys As String
  Dim stritems As String
  Dim Dic As Dictionary
  
  'オブジェクトの生成
  Set Dic = New Dictionary
  
  strKeys = ""
  stritems = ""
  
  'MakeDictionary :Input Keys and items
  If MakeDictionary(Dic) = False Then
  
    Debug.Print "MakeDictionary is failure."
    
    MsgBox "MakeDictionary is failure.", vbOKOnly, "failure"
  
  End If
  
  'Output MakeDictionary for FileList sheet
  If OutputDictionary(Dic) = False Then
  
    Debug.Print "OutputDictionary is failure."
    
    MsgBox "OutputDictionary is failure.", vbOKOnly, "failure"
  
  End If
  
  Set Dic = Nothing

  ThisWorkbook.Save
  
  MsgBox "Done!!"

  Call m_common.マクロ終了
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName:  sampleDictionary" & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "sampleDictionary : Failure"

    'Clear error
    Err.Clear
  
    Call m_common.マクロ終了
  
  End If

End Sub

