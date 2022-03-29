Attribute VB_Name = "m_main"
Option Explicit

'******************************************************************************
'FunctionName:DeleteIECookie
'Specifications�FClear Internet Explorer cookie information
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************

Sub DeleteIECookie()

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
   
  '��IE���N�����\��
  Set objIE = CreateObject("InternetExplorer.Application")
  
  objIE.Visible = True
   
  '�C���^�[�l�b�g�ꎞ�t�@�C�������Web�T�C�g�̃t�@�C�����폜
  'Delete temporary Internet and Web site files
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8"
  
  '�N�b�L�[��Web�T�C�g�̃f�[�^���폜
  'Delete cookies and website data
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2"
  
  '�������폜 Remove history
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1"
  
  '�t�H�[���f�[�^���폜���� Delete form data
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16"
  
  '�p�X���[�h���폜����@Delete Password
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32"
  
  '��L�܂ł̑S�Ẵf�[�^���폜�@Delete all data up to the above
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255"
  
  '��L�܂ł̑S�Ẵf�[�^+�A�h�I���ɂ���Đݒ肳�ꂽ�����܂ߑS�č폜
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

