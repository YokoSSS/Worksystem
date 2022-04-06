Attribute VB_Name = "m_common"
Option Explicit

'******************************************************************************
'FunctionName:Macrostart
'Specifications：Stops unnecessary movements during Macrostart to increase
'processing speed.
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub Macrostart()
    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に
End Sub

'******************************************************************************
'FunctionName:Macroend
'Specifications：Restarting operations that had been stopped to increase
'processing speed to increase wasteful movements during Macrostart.
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub Macroend()
    Application.StatusBar = False 'ステータスバーを消す
    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始
End Sub

'******************************************************************************
'FunctionName:GETフォルダ
'Specifications：フォルダダイアログでフォルダをピックアップ
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
'Function GETフォルダ(ByRef strFD As String) As Boolean
Function GETフォルダ()
Dim strFD As String
  On Error GoTo Err_Trap
  
  GETフォルダ = False
  
  With Application.FileDialog(msoFileDialogFolderPicker)
      
    If .Show = True Then
        
      strFD = .SelectedItems(1)
    
    Else

      MsgBox "キャンセルしました。処理を終了します。", vbInformation, "処理終了"

      Exit Function

    End If
  
  End With

  GETフォルダ = True

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
'FunctionName:GETファイル
'Specifications：ファイルダイアログでファイルをピックアップ
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Function GETファイル(ByRef strFL As String) As Boolean

  On Error GoTo Err_Trap

  GETファイル = False
  
  With Application.FileDialog(msoFileDialogFilePicker)
    
    If .Show = True Then
        
      strFL = .SelectedItems(1)
       
      Debug.Print "GETファイル名:　" & strFL

    Else
      
      MsgBox "キャンセルしました。処理を終了します。", vbInformation, "処理終了"
    
    End If

  End With

  GETファイル = True

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
'FunctionName:指定したパスでフォルダ生成
'Specifications：指定したパスでフォルダ生成
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Function 指定したパスでフォルダ生成(ByRef sFdPath As String) As Boolean
    
  Dim FSO As Object
  
  On Error GoTo Err_Trap
  
  指定したパスでフォルダ生成 = False
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FSO.CreateFolder sFdPath
  
  Set FSO = Nothing

  指定したパスでフォルダ生成 = True
    
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
'FunctionName:フォルダの存在有無
'Specifications：フォルダが存在するかどうかを調べるFunctionプロシージャ
'Arguments：nothing
'ReturnValue:nothing
'Note：
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
