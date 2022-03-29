Attribute VB_Name = "m_main"
Option Explicit

'******************************************************************************
'FunctionName:DeleteIECookie
'Specifications：Clear Internet Explorer cookie information
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************

Sub DeleteIECookie()

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
   
  '■IEを起動→表示
  Set objIE = CreateObject("InternetExplorer.Application")
  
  objIE.Visible = True
   
  'インターネット一時ファイルおよびWebサイトのファイルを削除
  'Delete temporary Internet and Web site files
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8"
  
  'クッキーとWebサイトのデータを削除
  'Delete cookies and website data
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2"
  
  '履歴を削除 Remove history
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1"
  
  'フォームデータを削除する Delete form data
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16"
  
  'パスワードを削除する　Delete Password
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32"
  
  '上記までの全てのデータを削除　Delete all data up to the above
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255"
  
  '上記までの全てのデータ+アドオンによって設定された情報も含め全て削除
  'Delete all data up to the above plus any information set by add-ons
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351"
  
  objIE.quite
  
  Set objIE = Nothing

  ThisWorkbook.Save
  
  MsgBox "Done!!"

  Call m_common.MacroEnd
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "DeleteIECookie: " & Err.Number & " " & Err.Description
    MsgBox "FunctionName DeleteIECookie: " & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "DeleteIECookie: Error"

    'Clear error
    Err.Clear
  
    Call m_common.MacroEnd
  
  End If

End Sub

