Attribute VB_Name = "m_work"
Option Explicit
'******************************************************************************
'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'''-----32ビット用-----
'Sleep機能を使うAPI
'Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
'Public Declare Sub Sleep Lib "KERNEL32" (ByVal ms As Long)
'Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
'   (ByVal pCaller As Long, _
'    ByVal szURL As String, _
'    ByVal szFileName As String, _
'    ByVal dwReserved As Long, _
'    ByVal lpfnCB As Long) As Long

'-----64ビット用-----
'GETTICK
Private Declare PtrSafe Function GetTickCount Lib "user32" () As Long
'強制的に最前面にする
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
  (ByVal hWnd As Long) As Long
'最小化されているか調べる
Private Declare PtrSafe Function IsIconic Lib "user32" _
  (ByVal hWnd As Long) As Long
'元の大きさに戻すAPI
Private Declare PtrSafe Function ShowWindowAsync Lib "user32" _
  (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'ダイアログが表示されたか判定する
Private Declare PtrSafe Function GetLAstActivePopup Lib "user32" _
 (ByVal hWnd As Long) As Long

'//Sleep機能を使うAPI
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
'Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
   (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

'******************************************************************************
'FunctionName:GetTickCount_sample2
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub GetTickCount_sample2()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart


  Dim Starttime As Long

  Starttime = GetTickCount

  Do While GetTickCount - Starttime < 5000

    DoEvents

  Loop

  MsgBox "5秒経過しました。"

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: GetTickCount_sample2" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:すでに起動しているShellを取得する
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub すでに起動しているShellを取得する()
  
  '起動中のShellを格納する変数
  Dim colSh As Object
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
    
  '現在開いているIEとエクスプローラーをcolShに格納
  Set colSh = CreateObject("Shell.Application")

  '変数colShには複数のオブジェクトが格納されています。格納されたオブジェクトの数
  '(起動しているIEとエクスプローラーの数)は次の様に取得できます。
  MsgBox "格納されたオブジェクトの数: " & colSh.Windows.Count

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: すでに起動しているShellを取得する" & vbCrLf & _
    Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:OpenIE
'Specifications：IEを起動する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub OpenIE()

  Dim ie As Object

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
  
  Set ie = CreateObject("InternetExplorer.Application")

  ie.Visible = True

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: OpenIE" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:OpenURL
'Specifications：URLを指定してWebページへ移動する（動的ページ）
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub openURL()

  Dim ie As InternetExplorer

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
    
  Set ie = CreateObject("InternetExplorer.Application")

  ie.Visible = True

  '指定したURLに移動する
  ie.navigate "http://search.yahoo.co.jp/search?p=" & ActiveCell.Value

  ie.Quit

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: OpenURL" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:WaitTest
'Specifications：IEの「Busy」プロパティを監視する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub WaitTest()

  Dim ie As InternetExplorer

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Set ie = CreateObject("InternetExplorer.Application")

  ie.Visible = True

  '指定したURLに移動する
  ie.navigate "http://search.yahoo.co.jp/search?p=" & ActiveCell.Value

  Do While ie.Busy

    Debug.Print ie.Busy

    DoEvents

  Loop

  ie.Quit
  
  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: WaitTest" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:WaitTest2
'Specifications：IEの「ReadyState」プロパティを監視する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub WaitTest2()

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
   
  Dim ie As InternetExplorer

  Set ie = CreateObject("InternetExplorer.Application")

  ie.Visible = True

  '指定したURLに移動する
  ie.navigate "http://search.yahoo.co.jp/search?p=" & ActiveCell.Value

  Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE

    Debug.Print ie.Busy & ":" & ie.readyState

    DoEvents

  Loop

  MsgBox ie.document.body.innerText

  ie.Quit
  
  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: WaitTest2" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If
  
End Sub

'******************************************************************************
'FunctionName:SearchIEI
'Specifications：Webページのタイトルをして指定してIEを取得する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub SearchIEI()

  Dim colSh As Object
  Dim win As Object
  Dim strTemp As String
  Dim objIE As Object
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
  
  '現在開いているIEとエクスプローラーをcolShに格納する
  Set colSh = CreateObject("Shell.Application")
  
  'colShからWindowsを1つずつ取り出す
  For Each win In colSh.Windows

    'HTMLDocumentだったら
    If InStr(win.document, "HTMLDocument") > 0 Then
      
      'タイトルバーにPC Watchが含まれるか判定
      If InStr(win.document.Title, "PC Watch") > 0 Then
      
        '実数objIEに取得したwinを格納
        Set objIE = win

        'ループを抜ける
        Exit For

      End If

    End If

  Next

  If objIE Is Nothing Then
  
    MsgBox "探しているIEはありませんでした"
  
  Else
  

    'タイトルを表示する
    MsgBox objIE.document.Title & "がありました"
    
    objIE.Quit

  End If

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: SearchIEI" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:SearchIEI3
'Specifications：様々な条件で目的のIEを取得する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub SearchIEI3()

  Dim colSh As Object
  Dim win As Object
  Dim strTemp As String
  Dim objIE As Object
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
  
  '現在開いているIEとエクスプローラーをcolShに格納する
  Set colSh = CreateObject("Shell.Application")
  
  'colShからWindowsを1つずつ取り出す
  For Each win In colSh.Windows
    
    strTemp = ""
    
    On Error Resume Next

    strTemp = win.document.body.innerText

    On Error GoTo 0

    'タイトルバーにPC Watchが含まれるか判定
    If InStr(strTemp, "アップデート情報") > 0 Then
    
      '実数objIEに取得したwinを格納
      Set objIE = win

      'ループを抜ける
      Exit For

    End If
  
  Next

  If objIE Is Nothing Then
  
    MsgBox "探しているIEはありませんでした"
  
  Else

    'タイトルを表示する
    MsgBox objIE.document.Title & "がありました"
  
    objIE.Quit

  End If

  ThisWorkbook.Save

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: SearchIEI3" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:SearchIEI4
'Specifications：IEかどうかの判断をせずにタイトルで判定する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub SearchIEI4()

  Dim colSh As Object
  Dim win As Object
  Dim strTemp As String
  Dim objIE As Object
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
   
  '現在開いているIEとエクスプローラーをcolShに格納する
  Set colSh = CreateObject("Shell.Application")
  
  'colShからWindowsを1つずつ取り出す
  For Each win In colSh.Windows
    
    strTemp = ""
    
    On Error Resume Next
    strTemp = win.document.Title
    On Error GoTo 0

    'タイトルバーにPC Watchが含まれるか判定
    If InStr(strTemp, "PC Watch") > 0 Then
      
      '実数objIEに取得したwinを格納
      Set objIE = win

      'ループを抜ける
      Exit For

    End If
  
  Next

  If objIE Is Nothing Then
  
    MsgBox "探しているIEはありませんでした"
  
  Else

    'タイトルを表示する
    MsgBox objIE.document.Title & "がありました"
    
    objIE.Quit
  
  End If

  ThisWorkbook.Save

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: SearchIEI4" & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:最前面SetForegroundWindow
'Specifications：取得したIEを最前面に表示する
'Arguments：objIE/Object
'ReturnValue:nothing
'Note：
'&H0:SW_HIDE ウインドウを非表示にし、他のウインドウをアクティブにする
'&H2:SW_MAXIMIZE　ウインドウを最大化する
'&H3:SW_MINIMIZE　ウインドウを最小化する
'&H9:SW_RESTORE　最小化または最大化されていたウインドウを元の位置とサイズに戻す
'******************************************************************************
Function 最前面SetForegroundWindow(objIE As Object)

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
  
  '指定されたウインドウを最前面化する

  '最小化されている場合は元の大きさに戻す
  If IsIconic(objIE.hWnd) Then
    
    '9＝RESTORE：最小化前の状態
    ShowWindowAsync objIE.hWnd, &H9
    
  End If
  
  'IEを最前面に表示
  SetForegroundWindow (objIE.hWnd)

  objIE.Quit
  
  ThisWorkbook.Save
  
  Call m_common.Macroend
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: 最前面SetForegroundWindow" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Function


'******************************************************************************
'FunctionName:sendKey
'Specifications：nothing
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub sampleSendKey()

  'SendKeys string(,wait)
  Sendkeys "abc"

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: sampleSendKey" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:LastPopup
'Specifications：LastPopup
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub LastPopup()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.hWnd = GetLAstActivePopup(objIE.hWnd)

    DoEvents

  Loop

 'メッセージを表示する
  MsgBox "ダイアログが表示された。"

  objIE.Quit
  
  ThisWorkbook.Save

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: LastPopup" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Sendkeys1
'Specifications：実際にサイト上のポップアップウインドウを閉じる
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub Sendkeys1()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/IE.html"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.Busy

    Sleep 100
    
    Sendkeys "{ENTER}", True

  Loop

 'メッセージを表示する
  MsgBox "Enterキーが押下された。"
  
  objIE.Quit
  
  ThisWorkbook.Save
  
  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: Sendkeys1" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "Sendkeys1：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Sendkeys2
'Specifications：ファイルをダウンロードする。Navigate先にダウンロードするファイ
'ルを直接指定すると「ファイルの保存」ダイアログが表示されます。このダイアログ
'を閉じるサンプルです。
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub Sendkeys2()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"

  '3秒休んでからAlt+Sを送信
  Sleep 3000
    
  Sendkeys "%S", True

 'メッセージを表示する
  MsgBox "ファイルが保存された。"
  
  objIE.Quit
  
  ThisWorkbook.Save

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: Sendkeys2" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "Sendkeys2：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Sendkeys2_2
'Specifications：ファイルをダウンロードする。Navigate先にダウンロードするファイ
'ルを直接指定すると「ファイルの保存」ダイアログが表示されます。このダイアログ
'を閉じるサンプルです。
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub Sendkeys2_2()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.hWnd = GetLAstActivePopup(objIE.hWnd)

    DoEvents

  Loop
    
  Sendkeys "%S", True

 'メッセージを表示する
  MsgBox "ファイルが保存された。"
  
  objIE.Quit
  
  ThisWorkbook.Save

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: Sendkeys2_2" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "Sendkeys2_2：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Checkbusy
'Specifications：ページを開いた１秒後にポップアップウインドウとしてファイルの
'保存ダイアログが表示されるようになっていますが、Busyプロパティの変化はイミディ
'エイトウインドウにて確認することが出来ます。
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub Checkbusy()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.hWnd = GetLAstActivePopup(objIE.hWnd)

    DoEvents

    '3秒休む（ポップアップウインドウの表示時間に合わせて調整）
    Sleep 3000

  Loop

 'メッセージを表示する
  MsgBox "ファイルが保存された。kkk"
  
  objIE.Quit
  
  ThisWorkbook.Save

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: Checkbusy" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "Checkbusy：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Sendkeys3
'Specifications：ファイルをダウンロードする。Navigate先にダウンロードするファイ
'ルを直接指定すると「ファイルの保存」ダイアログが表示されます。このダイアログ
'を閉じるサンプルです。
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub Sendkeys3()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.hWnd = GetLAstActivePopup(objIE.hWnd)

    DoEvents

  Loop
    
  Sendkeys "%S", True

 'メッセージを表示する
  MsgBox "ファイルが保存された。kkk"
  
  objIE.Quit
  
  ThisWorkbook.Save

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: Sendkeys3" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "Sendkeys3：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Sendkeys3_2
'Specifications：ファイルをダウンロードする。Navigate先にダウンロードするファイ
'ルを直接指定すると「ファイルの保存」ダイアログが表示されます。このダイアログ
'を閉じるサンプルです。
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub Sendkeys3_2()


  On Error GoTo Err_Trap

  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")

  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"

  'busyの間待機
  Do While objIE.Busy

    Sleep 1

  Loop

  'busyとなるまで待機
  Do Until objIE.Busy

    Sleep 1

  Loop

  
  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.Busy

    DoEvents

  Loop

  Sendkeys "%S", True

 'メッセージを表示する
  MsgBox "ファイルが保存された。kkk"

  objIE.Quit

  ThisWorkbook.Save

  Call m_common.Macroend

  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: Sendkeys3_2" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "Sendkeys3_2：Error"

    'Clear error
    Err.Clear

    Call m_common.Macroend

  End If

End Sub

'******************************************************************************
'FunctionName:ShowBars
'Specifications：IEの表示を制御する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub ShowBars()

  On Error GoTo Err_Trap

  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")

  '可視化
  objIE.Visible = True
  
  '
  objIE.Toolbar = True
  
  '
  objIE.AddressBar = True
  
  '
  objIE.MenuBar = True
    
  '
  objIE.StatusBar = True
  
 'メッセージを表示する
  MsgBox "各種バーが表示された。"

  objIE.Quit

  ThisWorkbook.Save

  Call m_common.Macroend

  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: ShowBars" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "ShowBars：Error"

    'Clear error
    Err.Clear

    Call m_common.Macroend

  End If

End Sub

'******************************************************************************
'FunctionName:ChangeSizeAndLocation
'Specifications：ウインドウのサイズと位置を指定する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub ChangeSizeAndLocation()

  On Error GoTo Err_Trap

  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")

  '可視化
  objIE.Visible = True
  
  '
  objIE.Width = 800
  
  '
  objIE.Height = 600
  
  '
  objIE.Left = 100
    
  '
  objIE.Top = 0
  
  '
  objIE.resizable = True
 
 'メッセージを表示する
  MsgBox "ウインドウのサイズと位置が指定された。"

  objIE.Quit

  ThisWorkbook.Save

  Call m_common.Macroend

  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: ChangeSizeAndLocation" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "ChangeSizeAndLocation：Error"

    'Clear error
    Err.Clear

    Call m_common.Macroend

  End If

End Sub

'******************************************************************************
'FunctionName:ExecInvisible
'Specifications：ウインドウを表示せず処理を実行
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub ExecInvisible()

  On Error GoTo Err_Trap

  Call m_common.MacroStart

  Dim objIE As Object

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")

  '可視化オフ
  objIE.Visible = False
  
  '指定したURLに移動する
  objIE.navigate "http://yahoo.co.jp/"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

 
 'メッセージを表示する
  MsgBox objIE.document.Title

  objIE.Quit

  ThisWorkbook.Save

  Call m_common.Macroend

  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: ExecInvisible" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "ExecInvisible：Error"

    'Clear error
    Err.Clear

    Call m_common.Macroend

  End If

End Sub

'******************************************************************************
'FunctionName:useAnchor
'Specifications：IE画面上のハイパーリンクを使ってWebページを移動する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub useAnchor()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim anchor As HTMLAnchorElement

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-7.html"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

  'リンクの設定された文字列から処理対象を検索する
  For Each anchor In objIE.document.getElementsByTagName("A")

    If anchor.innerText = "やきそばパン vs 揚げパン" Then
      
      'ハイパーリンクをクリック
      anchor.Click

      'メッセージを表示する
      MsgBox "やきそばパン vs 揚げパンページが表示された。"
      
      Exit For
  
  End If

  Next anchor
  
  
  objIE.Quit
  
  ThisWorkbook.Save
  
  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: useAnchor" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "useAnchor：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:useButton
'Specifications：ボタンを操作する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub useButton()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim button As HTMLButtonElement

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-8.html"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

  'ボタン表面の文字列から処理対象を検索する
  For Each button In objIE.document.getElementsByTagName("INPUT")
    
    If button.Type = "button" And button.Value = "ボタン２" Then
      
      'ボタンをクリック
      button.Click

     'メッセージを表示する
      MsgBox "ボタンがクリックされた。"
      
      Exit For
  
    End If

  Next button

  objIE.Quit
    
  ThisWorkbook.Save
  
  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: useButton" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "useButton：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

''******************************************************************************
''FunctionName:openURL
''Specifications：IE画面上の画像を切り替える。
''Arguments：nothing
''ReturnValue:nothing
''Note：
''******************************************************************************
'Sub openURL1()
'
'  On Error GoTo Err_Trap
'
'  Call m_common.MacroStart
'
'  Dim objIE As Object
'  Dim Doc As HTMLDocument
'  Dim ObjTag As Object
'  Dim i As Long
'  Const Src1 As String = "button_01.png"
'  Const Src2 As String = "button_02.png"
'
'  'IEを開いてファイルの保存URLを開く
'  Set objIE = CreateObject("InternetExplorer.Application")
'  '可視化
'  objIE.Visible = True
'
'  '指定したURLに移動する
'  objIE.Navigate "http://book.impress.co.jp/appended/3384/4-9.html"
'
''  Call waitNavigation(objIE)
'
'  Set Doc = objIE.document
'
'  For i = 1 To 10
'
'    For Each ObjTag In Doc.getElementsByTagName("INPUT")
'
'      With ObjTag
'
'        On Error Resume Next
'
'        If InStr(.src, Srx1) > 0 Then
'
'          .src = Src2
'
'          '0.2秒停止後、画面を元に戻し、再度0.2秒停止
'          Sleep 200
'
'          .src = Src1
'
'          '0.2秒停止後、画面を元に戻し、再度0.2秒停止
'          Sleep 200
'
'          Exit For
'
'        End If
'
'        On Error GoTo 0
'
'      End With
'
'    Next ObjTag
'
'  Next i
'
'  objIE.Quit
'
'  ThisWorkbook.Save
'
'  Call m_common.MacroEnd
'
'  Exit Sub
'
'Err_Trap:
'
'  'When an error occurs, display the contents of the error in a message box.
'  If Err.Number <> 0 Then
'    '
'    Debug.Print Err.Number & " " & Err.Description
'    MsgBox "FunctionName: openURL" & vbCrLf & Err.Number & " " _
'    & Err.Description, vbOKOnly, "openURL：Error"
'
'    'Clear error
'    Err.Clear
'
'    Call m_common.MacroEnd
'
'  End If
'
'End Sub

'******************************************************************************
'FunctionName:GetTable3
'Specifications：タグの一覧表を作成する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub GetTable3()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim button As HTMLButtonElement

  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-10_3.html"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

  Call MakeList(objIE)
  Call MakeList2(objIE)
  Call MakeList3(objIE)
  Call MakeList4(objIE)
  Call MakeList5(objIE)


  objIE.Quit
    
  ThisWorkbook.Save
  
  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: GetTable3" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "GetTable3：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:MakeList
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub MakeList(objIE As InternetExplorer)
  
  Dim n As Long 'タグの通し番号
  Dim r As Long 'td,thタグの通し番号
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  'タグ格納用
  Dim wslist As Worksheet

  On Error GoTo Err_Trap
  
  Debug.Print "(Function: MakeList START)"
  '変数初期化
  n = 0
  r = 0
  
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/標準"
    
    Set Doc = objIE.document
    
  End With
    Debug.Print Doc.all.Length - 1
  'ボタン表面の文字列から処理対象を検索する
  For n = 0 To Doc.all.Length - 1
  
    With Doc.all(n)
    
'      Debug.Print "(tagName = " & .tagName & ")"
        If .tagName = "INPUT" Or .tagName = "TEXTAREA" Or .tagName = "SELECT" Or _
        .tagName = "A" Or .tagName = "DIV" Or .tagName = "SCRIPT" Or _
        .tagName = "TD" Or .tagName = "P" Or .tagName = "TR" Or _
        .tagName = "SPAN" Or .tagName = "STRONG" Or .tagName = "BR" Or .tagName = "TABLE" Or _
        .tagName = "TBODY" Or .tagName = "IMG" Or .tagName = "OPTION" Or .tagName = "CENTER" Or _
        .tagName = "HEAD" Or .tagName = "BODY" Or .tagName = "LABEL" Or .tagName = "LI" Or _
        .tagName = "UR" Or .tagName = "fieldset" Or .tagName = "form" Or .tagName = "H1" Or _
        .tagName = "H2" Or .tagName = "H3" Or .tagName = "H4" Or .tagName = "H5" Or _
        .tagName = "IFRAME" Or .tagName = "THEAD" Or .tagName = "BODY" Or _
        .tagName = "LEFT" Or .tagName = "RIGHT" Or .tagName = "HTML" Then
        'If .tagName = "TD" Or .tagName = "TH" Then
      
        r = r + 1
        
        '項目
        If r = 1 Then
        
          'number
          wslist.Cells(r, 1) = "Number"
          
          'タグの名前
          wslist.Cells(r, 2) = "タグの名前"
        
          'タグの通し番号
          wslist.Cells(r, 3) = "タグの通し番号"
        
          'td,thタグの通し番号
          wslist.Cells(r, 4) = "td,thタグの通し番号"
        
          'テキスト(SOURCEINDEX)
          wslist.Cells(r, 5) = "SOURCEINDEX"
          
          'テキスト(innertext)
          wslist.Cells(r, 6) = "テキスト(innertext)"
          
          'テキスト(outertext)
          wslist.Cells(r, 7) = "テキスト(outertext)"
      
          'HTML(outerhtml)
          wslist.Cells(r, 8) = "HTML(outerhtml)"
      
        '取得情報
        ElseIf r > 1 Then
          'number
          wslist.Cells(r, 1) = r - 1
          
          'タグの名前
          wslist.Cells(r, 2) = .tagName
        
          'タグの通し番号
          wslist.Cells(r, 3) = n
        
          'td,thタグの通し番号
          wslist.Cells(r, 4) = r
        
          'SOURCEINDEX
          wslist.Cells(r, 5) = .sourceIndex
        
          'テキスト(innertext)
          wslist.Cells(r, 6) = .innerText
          
          'テキスト(outertext)
          wslist.Cells(r, 7) = .outerText
      
          'HTML(outerhtml)
          wslist.Cells(r, 8) = .outerHTML
        
        End If
        
      End If

    End With

  Next n
  
  Set wslist = Nothing
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: GetTable3" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "MakeList2：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:MakeList2
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：537ばから855番目までのタグを調整する
'tdタグなら取得し、1列ずつ右にずらしながらセルに書き込む
'16個取得したら次の行の1列目に移る。
'******************************************************************************
Sub MakeList2(objIE As InternetExplorer)
  
  Dim n As Long 'タグの通し番号
  Dim r As Long 'td,thタグの通し番号
  Dim i As Long '
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  'タグ格納用
  Dim wslist As Worksheet

  On Error GoTo Err_Trap
  
  '変数初期化
  n = 0
  r = 0
  i = 0
  
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/標準"
    
    Set Doc = objIE.document
    
  End With
  
  '
  For i = 537 To 855
    
    If Doc.all(n).tagName = "TD" Then

      n = n + 1
      
      wslist.Cells(Int((n - 1) / 16) + 1, (n - 1) Mod 16 + 1) = _
        Doc.all(i).innerText
      
    End If
  
  Next i
    
  Set wslist = Nothing
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: GetTable3" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "MakeList2：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:MakeList3
'Specifications：テーブルが可変の場合の取得方法
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub MakeList3(objIE As InternetExplorer)
  
  Dim n As Long 'タグの通し番号
  Dim r As Long 'td,thタグの通し番号
  Dim i As Long '
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  'タグ格納用
  Dim wslist As Worksheet
  Dim StartTag As Long
  Dim FinishTag As Long
  
  On Error GoTo Err_Trap
  
  '変数初期化
  n = 0
  r = 0
  i = 0
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/標準"
    
    Set Doc = objIE.document
    
  End With
  
  'ドキュメント構成タグを１つずつ調査
  For i = 0 To Doc.all.Length - 1
    
    'thタグなら
    If Doc.all(i).tagName = "TH" Then
    
      If Doc.all(i).innerText = "液晶" Then
      
        StartTag = 1
        
        Exit For
        
      End If
      
    End If
      
  Next i
    
  'ドキュメント構成タグを１つずつ調査
  For i = StartTag To Doc.all.Length - 1
    
    'thタグなら
    If Doc.all(i).tagName = "TH" Then
    
      If Doc.all(i).innerText = "メーカー" Then
        
        FinishTag = 1

        Exit For
        
      End If
      
    End If
      
  Next i
  
  'ドキュメント構成タグを１つずつ調査
  For i = StartTag To FinishTag
    
    'tdタグなら
    If Doc.all(i).tagName = "TD" Then
    
      n = n + 1
    
      wslist.Cells(Int((n - 1) / 16) + 1, (n - 1) Mod 16 + 1) = _
        Doc.all(i).innerText
    End If
      
  Next i
        
  Set wslist = Nothing
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: GetTable3" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "MakeList3：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub


'******************************************************************************
'FunctionName:MakeList4
'Specifications：より高速に処理する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub MakeList4(objIE As InternetExplorer)
  
  Dim n As Long 'タグの通し番号
  Dim r As Long 'td,thタグの通し番号
  Dim i As Long '
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  'タグ格納用
  Dim ObjTD As Object  'タグ格納用
  Dim wslist As Worksheet

  On Error GoTo Err_Trap
  
  '変数初期化
  n = 0
  r = 0
  i = 0
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/標準"
    
   Set Doc = objIE.document
    
  End With
  
  For Each ObjTag In ObjTD
    
    n = n + 1
    
    wslist.Cells(n, 1) = n
        
    wslist.Cells(n, 1) = ObjTag.tagName
        
    wslist.Cells(n, 1) = ObjTag.innerText
      
  Next ObjTag
      
  Set wslist = Nothing
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: GetTable3" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "MakeList4：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub


'******************************************************************************
'FunctionName:MakeList5
'Specifications：すべてのテーブルの一覧表を作成するコードを作る
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub MakeList5(objIE As InternetExplorer)
  
  Dim n As Long 'タグの通し番号
  Dim r As Long 'td,thタグの通し番号
  Dim c As Long 'カラム
  Dim i As Long '
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  'タグ格納用
  Dim ObjTD As Object  'タグ格納用
  Dim wslist As Worksheet

  On Error GoTo Err_Trap
  
  '変数初期化
  n = 0
  r = 0
  i = 0
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/標準"
    
   Set Doc = objIE.document
    
  End With
  
  For i = 0 To Doc.all.Length - 1

    'tdタグかthタグ
    If Doc.all(i).tagName = "TH" Or Doc.all(i).tagName = "TD" Then
  
      wslist.Cells(r, c) = Doc.all(i).innerText

      c = c + 1
    'trタグ
    ElseIf Doc.all(i).tagName = "TR" Then
    
      r = r + 1
      
      c = 1

    End If
  
  Next i

  Set wslist = Nothing
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: GetTable3" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "MakeList5：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:useScript
'Specifications：スクリプトを実行する。Webページにメッセージボックスを表示する｡
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub useScript()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim button As HTMLButtonElement
  Dim pwin As HTMLWindow2
  
  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-13.html"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

  Set pwin = objIE.document.parentWindow
  
  pwin.alert ("VBAからalertを実行")
  
  objIE.Quit
    
  ThisWorkbook.Save
  
  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: useScript" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "useScript：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:useScript2
'Specifications：スクリプト処理の完了を待たずにVBAの後続処理を実行する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub useScript2()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim button As HTMLButtonElement
  Dim pwin As HTMLWindow2
  
  'IEを開いてファイルの保存URLを開く
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '可視化
  objIE.Visible = True

  '指定したURLに移動する
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-13.html"

  'ファイルを開くダイアログが表示されるまでループ
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

  Set pwin = objIE.document.parentWindow
  
  pwin.execScript "showMessage('VBAからshowMessageを実行')"
  
  '**********************************************************************
  'OKボタンをクリックしてIEのメッセージを閉じる前に4が実行されるように
  'するには上記の処理を以下の様に処理する
  'pwin.setTimeout "showMessage('VBAからshowMessageを非同期実行')", 0
  '**********************************************************************
  
  MsgBox "VBAの後続処理"
  
  objIE.Quit
    
  ThisWorkbook.Save
  
  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: useScript2" & vbCrLf & Err.Number & " " _
    & Err.Description, vbOKOnly, "useScript2：Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName: getElementList
'Specifications：'HTMLソース全体を取得する
'Arguments：getHTMLString / String
'ReturnValue:nothing
'Note：
'******************************************************************************
Public Function getHTMLString(ByVal objIE As InternetExplorer) As String
  
  Dim htdoc As HTMLDocument
  Dim ret As String
  Dim elle As IHTMLElement
  
  Set htdoc = objIE.document
  
  'HTMLソース全体を取得する
  Set elle = htdoc.getElementByTagName("HTML")(0)
  
'  ret = htdoc.getElementbyTAgName("HTML")(0).outerHTML & vbCrLf
  ret = elle.outerHTML & vbCrLf

  Set htdoc = Nothing


  getHTMLString = ret

End Function


'******************************************************************************
'FunctionName: getElementList
'Specifications：
'Arguments：htdoc / HTMLDocument
'ReturnValue:nothing
'Note：
'******************************************************************************
Public Function getElementList(ByVal htdoc As HTMLDocument)
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim ret As String
  
  ret = "TAG:" & vbTab & "Type" & vbTab & "ID" & vbTab & "Name" & vbTab & vbCrLf
  
  Dim element As Object
  
  For Each element In htdoc.all
  
    Select Case UCase(element.tagName)
      
      'Evaluate whether the tag type is INPUT/TEXTAREA/SELECT.
      Case "INPUT", "TEXTAREA", "SELECT"
      
        ret = ret & element.tagName & vbTab & element.Type & vbTab & _
        element.ID & vbTab & element.Name & vbTab & element.Value & vbCrLf

    End Select
    
  Next element
  
  getElementList = ret

  ThisWorkbook.Save
  
  MsgBox "Done!!"

  Call m_common.Macroend
  
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
'FunctionName: printMain01
'Specifications：
'Arguments：htdoc / HTMLDocument
'ReturnValue:nothing
'Note：
'******************************************************************************
Public Sub printMain01(ByVal objIE As InternetExplorer)

  On Error GoTo Err_Trap

  Dim HTMLstring  As String
  Dim FileName    As String
  Dim FileNum     As Long

  'HTML全部取得
'  HTMLstring = getHTMLString(objIE)
  HTMLstring = getHTMLString(objIE)
  
  FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & ".txt"
  
  'ファイル番号の取得
  FileNum = FreeFile()
  
  Open FileName For Output As #FileNum
  
    Print #FileNum, HTMLstring

  Close #FileNum
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "printMain01: " & Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "printMain01: Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  End If

End Sub

'******************************************************************************
'FunctionName: getHTMLString
'Specifications：
'Arguments：htdoc / HTMLDocument
'ReturnValue:nothing
'Note：
'******************************************************************************
Public Function getHTMLString2(ByVal container As Object, ByVal objIE As Object, Optional depth As Long = 0) As String

  
  On Error GoTo Err_Trap
  
'  Call m_common.Macrostart
  Dim ErrorInfo As String
  Dim htdoc As HTMLDocument
  
  On Error Resume Next
  
'  Set htdoc = container.document
  Set htdoc = objIE.document

  If Err.Number <> 0 Then
  
    ErrorInfo = Trim(str(Err.Number)) & ":" & Err.Description
  
  End If
  
  On Error GoTo 0
  
  Dim ret As String
  
  'リターン値を区切り線と階層情報で初期化
  ret = "-------------------------------------------------------------" & vbCrLf
  
  ret = ret & "[" & Trim(str(depth)) & "階層]" & vbCrLf
  
  Dim i As Integer
  
  'If HTML can be retrieved
  If Not htdoc Is Nothing Then
  
    'Frame and document information (フレームと文書の情報)
'    ret = ret & htdoc.Title & " | " & htdoc.Location & " (" & container.Name & ")" & vbCrLf
    ret = ret & htdoc.Title & " | " & htdoc.Location & vbCrLf
    
    ret = ret & "-------------------------------------------------------------" & vbCrLf
  
'    'Obtain a list of screen components
'    ret = ret & htdoc.getElementList(htdoc) & vbCrLf


'    ret = ret & "-------------------------------------------------------------" & vbCrLf
    
    '(HTMLタグ要素(画面に一つ))
    ret = ret & objIE.document.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
    
  
    For i = 0 To objIE.document.frames.Length - 1
    
      'Recurses if frames are present (フレームがある場合は再帰する)
      ret = ret & htdoc.getHTMLString(htdoc.frames(i), depth + 1)
    
    Next i
  
  Else
  
    ret = ret & "-------------------------------------------------------------" & vbCrLf
    
    'Output error information if HTML could not be obtained
    '(HTMLが取得出来なかった場合はエラー情報を出力)
    ret = ret & ErrorInfo
    
  End If
  
  getHTMLString2 = ret

'  ThisWorkbook.Save
'
'  MsgBox "Done!!"
'
'  Call m_common.Macroend
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "getHTMLString2 : " & Err.Number & " " & Err.Description
    MsgBox "FunctionName: getHTMLString2 " & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "getHTMLString2 : Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Function

'******************************************************************************
'FunctionName: getHTMLFramesString
'Specifications：フレームを含めたすべてのHTMLを出力する
'Arguments：htdoc / HTMLDocument
'ReturnValue:nothing
'Note：
'******************************************************************************
Public Function getHTMLFramesString(ByVal container As Object, ByVal objIE As Object) As String
  
  On Error GoTo Err_Trap
  
'  Call m_common.Macrostart
  Dim ErrorInfo As String
  Dim htdoc As HTMLDocument
  Dim ret As String
  Dim i As Integer
  
'  Set htdoc = container.document
  Set htdoc = objIE.document

  'リターン値を区切り線と階層情報で初期化
  ret = "-------------------------------------------------------------" & vbCrLf
  
  '(HTMLタグ要素(画面に一つ))
  ret = ret & objIE.document.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
  
  
  For i = 0 To objIE.document.frames.Length - 1
    
    'Recurses if frames are present (フレームがある場合は再帰する)
    ret = ret & htdoc.getHTMLString(objIE.document.frames(i))
    
  Next i
  ret = ret & htdoc.getHTMLString(objIE.document.frames(i))

  getHTMLFramesString = ret

  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "getHTMLFramesString : " & Err.Number & " " & Err.Description
    MsgBox "FunctionName: getHTMLFramesString " & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "getHTMLFramesString : Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Function

'******************************************************************************
'FunctionName: getHTMLString3
'Specifications：エラーで調査が中断されないようにし、すべてのHTMLを出力する
'Arguments：htdoc / HTMLDocument
'ReturnValue:nothing
'Note：
'******************************************************************************
Public Function getHTMLString3(ByVal container As Object, ByVal objIE As Object) As String
  
  On Error GoTo Err_Trap
  
'  Call m_common.Macrostart
  Dim ErrorInfo As String
  Dim htdoc As HTMLDocument
  Dim ret As String
  Dim i As Integer
  
'  Set htdoc = objIE.container.document
  Set htdoc = objIE.document

  If Err.Number <> 0 Then
  
    ErrorInfo = Trim(str(Err.Number)) & ":" & Err.Description
  
  End If
  
  On Error GoTo 0

  'リターン値を区切り線と階層情報で初期化
  ret = "-------------------------------------------------------------" & vbCrLf
  
  ret = ret & objIE.document.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
  
  If Not htdoc Is Nothing Then
    
    '(HTMLタグ要素(画面に一つ))
    ret = ret & objIE.document.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
  
    For i = 0 To objIE.document.frames.Length - 1
      
      'Recurses if frames are present (フレームがある場合は再帰する)
      ret = ret & getHTMLString3(objIE.document.frames(i), objIE)
      
    Next i
  
  Else
  
    'HTMLが取得出来なかった場合はエラー情報を出力する
    ret = ret & ErrorInfo

  End If
  
  getHTMLString3 = ret

  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "getHTMLString3 : " & Err.Number & " " & Err.Description
    MsgBox "FunctionName: getHTMLString3 " & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "getHTMLString3 : Error"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Function
'【objIE.document.getElementsByTagName("h2")(0).outerHTML】
'objIE = InternetExplorerオブジェクト
'document = HTMLドキュメントのオブジェクト(Documentオブジェクト)
'getElementsByTagName("h2") = HTMLドキュメント内のすべてのh2要素(GetElementsByTagNameメソッド)
'getElementsByTagName("h2")(0) = h2要素コレクションの1番目のh2要素オブジェクト
'outerHTML = 1番目のh2要素オブジェクトの要素タグとその中に含まれるHTMLコード


'******************************************************************************
'FunctionName: getElementList2
'Specifications：
'Arguments：htdoc / HTMLDocument
'ReturnValue:nothing
'Note：
'******************************************************************************
Public Function getElementList2(ByVal htdoc As HTMLDocument)
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim ret As String
  Dim element As Object
  Dim i  As Long    '部品番号用変数を宣言する
  
  'ヘッダーの先頭に番号を示す項目を追加する
  ret = "#" & vbTab & "タグ" & vbTab & "Type" & vbTab & "ID" & vbTab & "Name" & _
    vbTab & "Value" & vbCrLf

  'ドキュメント全要素に対して処理
  For Each element In htdoc.all
  
    Select Case UCase(element.tagName)
      
      'Evaluate whether the tag type is INPUT/TEXTAREA/SELECT.
      Case "INPUT", "TEXTAREA", "SELECT"
        
        '画面ぶひんじょうほうの先頭に番号を記録する
        ret = ret & CStr(i) & vbTab & element.tagName & vbTab & element.Type & vbTab & _
        element.ID & vbTab & element.Name & vbTab & element.Value & vbCrLf

        '画面に番号を書き戻す
        If UCase(element.Type) <> "HIDDEN" Then
        
          element.outerHTML = element.outerHTML & "&nbsp;<bstyle=""color:blue;"">[" & CStr(i) & "]</b>"
        
        End If
        
        '番号をインクリメントする
        i = i + 1

    End Select
    
  Next element
  
  getElementList2 = ret

  ThisWorkbook.Save
  
  MsgBox "Done!!"

  Call m_common.Macroend
  
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
'FunctionName: getElementList3
'Specifications：
'Arguments：htdoc / HTMLDocument
'ReturnValue:nothing
'Note：
'******************************************************************************
Public Function getElementList3(ByVal htdoc As HTMLDocument)
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim ret As String
  Dim element As Object
  Dim i  As Long    '部品番号用変数を宣言する
  
  'ヘッダーの先頭に番号を示す項目を追加する
  ret = "#" & vbTab & "タグ" & vbTab & "Type" & vbTab & "ID" & vbTab & "Name" & _
    vbTab & "Value" & vbCrLf

  'ドキュメント全要素に対して処理
  For Each element In htdoc.all
  
    Select Case UCase(element.tagName)
      
      'Evaluate whether the tag type is INPUT/TEXTAREA/SELECT.
      Case "INPUT", "TEXTAREA", "SELECT", "A", "DIV", "SCRIPT", "TD", "P", "TR", _
        "SPAN", "STRONG", "BR", "TABLE", "TBODY", "IMG", "OPTION", "CENTER", "HEAD", _
        "BODY", "LABEL", "LI", "UR", "fieldset", "form", "H1", "H2", "H3", "H4", "H5", _
        "IFRAME", "THEAD", "BODY", "LEFT", "RIGHT", "HTML"
      
        '画面ぶひんじょうほうの先頭に番号を記録する
        ret = ret & CStr(i) & vbTab & element.tagName & vbTab & element.Type & vbTab & _
        element.ID & vbTab & element.Name & vbTab & element.Value & vbCrLf

        '画面に番号を書き戻す
        If UCase(element.Type) <> "HIDDEN" Then
        
          element.outerHTML = element.outerHTML & "&nbsp;<bstyle=""color:blue;"">[" & CStr(i) & "]</b>"
        
        Else
          
          element.outerHTML = element.outerHTML & "&nbsp;<bstyle=""color:blue;"">[" & CStr(i) & "]</b>"
        
        End If
        
        '番号をインクリメントする
        i = i + 1

    End Select
    
  Next element
  
  getElementList3 = ret

  ThisWorkbook.Save
  
  MsgBox "Done!!"

  Call m_common.Macroend
  
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



