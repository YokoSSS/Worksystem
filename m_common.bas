Attribute VB_Name = "m_common"
Option Explicit

'******************************************************************************
'FunctionName:MacroStart
'Specifications：MacroStart時、処理スピードを高める為に無駄な動きをする動作を
'停止させる
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************

Sub MacroStart()
    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に
End Sub

'******************************************************************************
'FunctionName:MacroEnd
'Specifications：MacroStart時、処理スピードを高める為に無駄な動きをする動作を
'停止させていたものを再稼働させる
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub MacroEnd()
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
    
    Call m_common.MacroEnd

    'Clear error
    Err.Clear
  
  End If

End Function

'******************************************************************************
'FunctionName:GETFile
'Specifications：Pick a file in the file dialog up.
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Function GETFile(ByRef strFL As String) As Boolean

  On Error GoTo Err_Trap

  GETFile = False
  
  With Application.FileDialog(msoFileDialogFilePicker)
    
    If .Show = True Then
        
      strFL = .SelectedItems(1)
       
      Debug.Print "GETFileName :　" & strFL

    Else
      
      MsgBox "Canceled. The process is terminated.", vbInformation, _
      "End of processing"
    
    End If

  End With

  GETFile = True

  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "GETFile: " & Err.Number & " " & Err.Description
    MsgBox "GETFile: " & Err.Number & " " & Err.Description, vbOKOnly, _
    "GETFile: Error"

    'Clear error
    Err.Clear

    Call m_common.MacroEnd
  
  End If

End Function

'******************************************************************************
'FunctionName:CreateFolder
'Specifications：Create a folder with the specified path
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Function CreateFolder(ByRef sFdPath As String) As Boolean
    
  Dim FSO As Object
  
  On Error GoTo Err_Trap
  
  CreateFolder = False
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FSO.CreateFolder sFdPath
  
  Set FSO = Nothing

  CreateFolder = True
    
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear

    Call m_common.MacroEnd
  
  End If

End Function

'******************************************************************************
'FunctionName:ExistenceOrNonexistenceFolders
'Specifications：Existence or non-existence of folders
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Function ExistenceOrNonexistenceFolders(folder_path As String) As Boolean
  
  On Error GoTo Err_Trap
  
  If Dir(folder_path, vbDirectory) = "" Then
    
    ExistenceOrNonexistenceFolders = False
  
  Else
    
    ExistenceOrNonexistenceFolders = True
  
  End If
  
  Exit Function

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print "ExistenceOrNonexistenceFolders: " & Err.Number & " " & _
    Err.Description
    MsgBox "ExistenceOrNonexistenceFolders : " & Err.Number & " " & _
    Err.Description, vbOKOnly, "ExistenceOrNonExistenceFolders: Error"

    'Clear error
    Err.Clear

    Call m_common.MacroEnd
  
  End If

End Function
