Attribute VB_Name = "m_main"
Option Explicit
'//Sleep機能を使うAPI
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)


'******************************************************************************
'FunctionName:
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************

Sub sample()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  ThisWorkbook.Save

  MsgBox "Done!!"

  Call m_common.Macroend
  
  Exit Sub

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
  
End Sub

'******************************************************************************
'FunctionName:Login
'Specifications：
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub Login()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IEオブジェクトを準備
Dim htmlDoc    As HTMLDocument     'HTMLドキュメントオブジェクトを準備
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElementオブジェクトを準備
Dim eltext     As IHTMLElement '使用していない
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'デスクトップに"WORK"フォルダを作成する
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****ログイン*****
  'デスクトップにフォルダパスを拾う
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  'エイジスリサーチ様の
  IDusername = "1124_senoo"   'ログインユーザーネーム
  IDpassword = "to4lklp7"     'ログインパスワード

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IEオブジェクトをセットする
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IEを表示
    objIE.Visible = True

    'IEでURLを開く
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '読み込み待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIEで読み込まれているHTMLドキュメントをセット
    Set htmlDoc = objIE.document
    'IDをセット
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    
  '  Call printMain01(objIE)
  
  
  '    objIE.container
      
  '  str = getHTMLString(objIE)
  '
  '  Debug.Print str
  '
  '
    i = 0
    '
    str = getHTMLString2(objIE.container, objIE, i)
  
  
    Debug.Print str
    
    FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & "_" & j - 1 & ".txt"
    
    'ファイル番号の取得
    FileNum = FreeFile()
    
    Open FileName For Output As #FileNum
    
      Print #FileNum, str
  
    Close #FileNum
    
    objIE.Quit
    
  Next j

goal:
      
  If Not objIE Is Nothing = True Then objIE.Quit


  Set wsh = Nothing
  Set objIE = Nothing
  Set htmlDoc = Nothing
  Set elFormID = Nothing
  Set elFormpass = Nothing
  Set elbutton = Nothing

  ThisWorkbook.Save

  MsgBox "Done!!"

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    GoTo goal
  
  End If

End Sub


'******************************************************************************
'FunctionName:SampleLogin002
'Specifications：フレームを含めたすべてのHTMLを出力する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub SampleLogin002()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IEオブジェクトを準備
Dim htmlDoc    As HTMLDocument     'HTMLドキュメントオブジェクトを準備
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElementオブジェクトを準備
Dim eltext     As IHTMLElement '使用していない
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'デスクトップに"WORK"フォルダを作成する
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****ログイン*****
  'デスクトップにフォルダパスを拾う
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  'エイジスリサーチ様の
  IDusername = "1124_senoo"   'ログインユーザーネーム
  IDpassword = "to4lklp7"     'ログインパスワード

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IEオブジェクトをセットする
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IEを表示
    objIE.Visible = True

    'IEでURLを開く
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '読み込み待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIEで読み込まれているHTMLドキュメントをセット
    Set htmlDoc = objIE.document
    'IDをセット
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    
  '  Call printMain01(objIE)
  
  
  '    objIE.container
      
  '  str = getHTMLString(objIE)
  '
  '  Debug.Print str
  '
  '
    i = 0
    '
    str = getHTMLFramesString(objIE.container, objIE)
  
  
    Debug.Print str
    
    FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & "_" & j - 1 & ".txt"
    
    'ファイル番号の取得
    FileNum = FreeFile()
    
    Open FileName For Output As #FileNum
    
      Print #FileNum, str
  
    Close #FileNum
    
    objIE.Quit
    
  Next j

goal:
      
  If Not objIE Is Nothing = True Then objIE.Quit


  Set wsh = Nothing
  Set objIE = Nothing
  Set htmlDoc = Nothing
  Set elFormID = Nothing
  Set elFormpass = Nothing
  Set elbutton = Nothing

  ThisWorkbook.Save

  MsgBox "Done!!"

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    GoTo goal
  
  End If

End Sub

'******************************************************************************
'FunctionName:getHTMLString3Login003
'Specifications：エラーで調査が中断されないようにし、すべてのHTMLを出力する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub getHTMLString3Login003()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IEオブジェクトを準備
Dim htmlDoc    As HTMLDocument     'HTMLドキュメントオブジェクトを準備
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElementオブジェクトを準備
Dim eltext     As IHTMLElement '使用していない
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'デスクトップに"WORK"フォルダを作成する
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****ログイン*****
  'デスクトップにフォルダパスを拾う
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  'エイジスリサーチ様の
  IDusername = "1124_senoo"   'ログインユーザーネーム
  IDpassword = "to4lklp7"     'ログインパスワード

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IEオブジェクトをセットする
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IEを表示
    objIE.Visible = True

    'IEでURLを開く
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '読み込み待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIEで読み込まれているHTMLドキュメントをセット
    Set htmlDoc = objIE.document
    'IDをセット
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    
  '  Call printMain01(objIE)
  
  
  '    objIE.container
      
  '  str = getHTMLString(objIE)
  '
  '  Debug.Print str
  '
  '
    i = 0
    '
    str = getHTMLString3(objIE.container, objIE)
  
  
    Debug.Print str
    
    FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & "_" & j - 1 & ".txt"
    
    'ファイル番号の取得
    FileNum = FreeFile()
    
    Open FileName For Output As #FileNum
    
      Print #FileNum, str
  
    Close #FileNum
    
    objIE.Quit
    
  Next j

goal:
      
  If Not objIE Is Nothing = True Then objIE.Quit


  Set wsh = Nothing
  Set objIE = Nothing
  Set htmlDoc = Nothing
  Set elFormID = Nothing
  Set elFormpass = Nothing
  Set elbutton = Nothing

  ThisWorkbook.Save

  MsgBox "Done!!"

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    GoTo goal
  
  End If

End Sub


'******************************************************************************
'FunctionName:getHTMLStringLogin004
'Specifications：画面と解析情報を照合するための番号を表示する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub getHTMLStringLogin004()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IEオブジェクトを準備
Dim htmlDoc    As HTMLDocument     'HTMLドキュメントオブジェクトを準備
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElementオブジェクトを準備
Dim eltext     As IHTMLElement '使用していない
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'デスクトップに"WORK"フォルダを作成する
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****ログイン*****
  'デスクトップにフォルダパスを拾う
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  'エイジスリサーチ様の
  IDusername = "1124_senoo"   'ログインユーザーネーム
  IDpassword = "to4lklp7"     'ログインパスワード

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IEオブジェクトをセットする
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IEを表示
    objIE.Visible = True

    'IEでURLを開く
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '読み込み待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIEで読み込まれているHTMLドキュメントをセット
    Set htmlDoc = objIE.document
    'IDをセット
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    
    i = 0
    '
    str = getElementList2(htmlDoc)
  
  
    Debug.Print str
    
    FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & "_" & j - 1 & ".txt"
    
    'ファイル番号の取得
    FileNum = FreeFile()
    
    Open FileName For Output As #FileNum
    
      Print #FileNum, str
  
    Close #FileNum
    
    objIE.Quit
    
  Next j

goal:
      
  If Not objIE Is Nothing = True Then objIE.Quit


  Set wsh = Nothing
  Set objIE = Nothing
  Set htmlDoc = Nothing
  Set elFormID = Nothing
  Set elFormpass = Nothing
  Set elbutton = Nothing

  ThisWorkbook.Save

  MsgBox "Done!!"

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    GoTo goal
  
  End If

End Sub

'******************************************************************************
'FunctionName:getHTMLStringLogin005
'Specifications：画面と解析情報を照合するための番号を表示する
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub getHTMLStringLogin005()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IEオブジェクトを準備
Dim htmlDoc    As HTMLDocument     'HTMLドキュメントオブジェクトを準備
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElementオブジェクトを準備
Dim eltext     As IHTMLElement '使用していない
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'デスクトップに"WORK"フォルダを作成する
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****ログイン*****
  'デスクトップにフォルダパスを拾う
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  'エイジスリサーチ様の
  IDusername = "1124_senoo"   'ログインユーザーネーム
  IDpassword = "to4lklp7"     'ログインパスワード

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IEオブジェクトをセットする
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IEを表示
    objIE.Visible = True

    'IEでURLを開く
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '読み込み待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIEで読み込まれているHTMLドキュメントをセット
    Set htmlDoc = objIE.document
    'IDをセット
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'IDを入力
    elFormID.Value = IDusername
    'Passを入力
    elFormpass.Value = IDpassword
    '送信
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
    Sleep 15000
    
    
    i = 0
    
    Call MakeList(objIE)
    
    '
    str = getElementList3(htmlDoc)
  
  
    Debug.Print str
    
    FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & "_" & j - 1 & ".txt"
    
    'ファイル番号の取得
    FileNum = FreeFile()
    
    Open FileName For Output As #FileNum
    
      Print #FileNum, str
  
    Close #FileNum
    
    objIE.Quit
    
  Next j

goal:
      
  If Not objIE Is Nothing = True Then objIE.Quit


  Set wsh = Nothing
  Set objIE = Nothing
  Set htmlDoc = Nothing
  Set elFormID = Nothing
  Set elFormpass = Nothing
  Set elbutton = Nothing

  ThisWorkbook.Save

  MsgBox "Done!!"

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
    GoTo goal
  
  End If

End Sub

