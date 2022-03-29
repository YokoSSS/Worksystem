Attribute VB_Name = "m_work"
Option Explicit


Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32.dll" _
(ByVal hwnd As Long, ByRef ProcessId As Long) As Long

'******************************************************************************
'FunctionName:GetWindowProcessId
'Specifications：Obtain a Windows system processID.
'Arguments：hwnd / Long
'ReturnValue: GetWindowProcessId / Long
'Note：
'vbHide  0 Windows that have focus and are hidden
'vbNormalFocus 1 A window that has focus and is restored to its original size
'and position
'vbMinimizedFocus  2 Window with focus and minimized display
'vbMaximizedFocus  3 Window with focus and maximized display
'vbNormalNoFocus 4 An unfocused window that is restored to the size and
'position it was in when the window was last closed. The currently active
'window remains active.
'vbMinimizedNoFocus  6 A window without focus that is displayed minimized.
'The currently active window remains active.
'******************************************************************************
Function GetWindowProcessId(ByVal hwnd As Long) As Long
  
  On Error GoTo Err_Trap
  
  Dim ProcessId As Long
  
  GetWindowThreadProcessId hwnd, ProcessId
  
  Debug.Print "ProcessId  OK　" & ProcessId&
  
  GetWindowProcessId = ProcessId
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "GetWindowProcessId: " & Err.Number & " " & Err.Description
    MsgBox "FunctionName:  GetWindowProcessId" & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "GetWindowProcessId : Error"

    'Clear error
    Err.Clear
  
    Call m_common.マクロ終了
  
  End If

End Function

'******************************************************************************
'FunctionName:GetExplorerwindows
'Specifications：Get Explorer process ID to minimize and control windows.
'Arguments：Nothing
'ReturnValue: Nothing
'Note：
'******************************************************************************
Sub GetExplorerwindows()
  
  On Error GoTo Err_Trap
  
  Dim Shell As Object
  Dim ProcessId As Long
  Dim ie As Object
  Set Shell = CreateObject("Shell.Application")
  
'  ProcessId = VBA.Shell("explorer.exe /separate", vbHide)
  ProcessId = VBA.Shell("explorer.exe /separate", vbMinimizedFocus)
' ProcessId = VBA.Shell("explorer.exe /separate", vbNormalFocus)
  Debug.Print ProcessId
  
  Do
    '立ち上がっているウインドウズアプリケーションを表示する
    For Each ie In Shell.Windows()
      
      If ie.Visible = True Then
        
        Debug.Print ie.Name & "ProcessId  ?????"
        
        '立ち上がっているアプリケーションが"エクスプローラー"だったら
        If ie.Name = "エクスプローラー" Then
          
          'プロセスIDを取得しにいく
          GetWindowProcessId (ie.hwnd)
            
          '処理を抜ける
          Exit Do
          
        End If
          
      ElseIf GetWindowProcessId(ie.hwnd) = ProcessId Then
        
        '処理を抜ける
        Exit Do
      
      End If
    
    Next
    
    Application.Wait [NOW()+"0:00:00.1"]
  
  Loop
  
'  ie.Visible = True
  ie.Visible = False

  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "GetExplorerwindows: " & Err.Number & " " & Err.Description
    MsgBox "FunctionName:  GetExplorerwindows" & vbCrLf & Err.Number & " " & _
    Err.Description, vbOKOnly, "GetExplorerwindows : Error"

    'Clear error
    Err.Clear
  
    Call m_common.マクロ終了
  
  End If

End Sub


