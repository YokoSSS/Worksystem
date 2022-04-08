Attribute VB_Name = "m_work"
Option Explicit
'******************************************************************************
'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'''-----32�r�b�g�p-----
'Sleep�@�\���g��API
'Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
'Public Declare Sub Sleep Lib "KERNEL32" (ByVal ms As Long)
'Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
'   (ByVal pCaller As Long, _
'    ByVal szURL As String, _
'    ByVal szFileName As String, _
'    ByVal dwReserved As Long, _
'    ByVal lpfnCB As Long) As Long

'-----64�r�b�g�p-----
'GETTICK
Private Declare PtrSafe Function GetTickCount Lib "user32" () As Long
'�����I�ɍőO�ʂɂ���
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
  (ByVal hWnd As Long) As Long
'�ŏ�������Ă��邩���ׂ�
Private Declare PtrSafe Function IsIconic Lib "user32" _
  (ByVal hWnd As Long) As Long
'���̑傫���ɖ߂�API
Private Declare PtrSafe Function ShowWindowAsync Lib "user32" _
  (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'�_�C�A���O���\�����ꂽ�����肷��
Private Declare PtrSafe Function GetLAstActivePopup Lib "user32" _
 (ByVal hWnd As Long) As Long

'//Sleep�@�\���g��API
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
'Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
   (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

'******************************************************************************
'FunctionName:GetTickCount_sample2
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub GetTickCount_sample2()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart


  Dim Starttime As Long

  Starttime = GetTickCount

  Do While GetTickCount - Starttime < 5000

    DoEvents

  Loop

  MsgBox "5�b�o�߂��܂����B"

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
'FunctionName:���łɋN�����Ă���Shell���擾����
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub ���łɋN�����Ă���Shell���擾����()
  
  '�N������Shell���i�[����ϐ�
  Dim colSh As Object
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
    
  '���݊J���Ă���IE�ƃG�N�X�v���[���[��colSh�Ɋi�[
  Set colSh = CreateObject("Shell.Application")

  '�ϐ�colSh�ɂ͕����̃I�u�W�F�N�g���i�[����Ă��܂��B�i�[���ꂽ�I�u�W�F�N�g�̐�
  '(�N�����Ă���IE�ƃG�N�X�v���[���[�̐�)�͎��̗l�Ɏ擾�ł��܂��B
  MsgBox "�i�[���ꂽ�I�u�W�F�N�g�̐�: " & colSh.Windows.Count

  Call m_common.Macroend
  
  Exit Sub

Err_Trap:

  'When an error occurs, display the contents of the error in a message box.
  If Err.Number <> 0 Then
    '
    Debug.Print Err.Number & " " & Err.Description
    MsgBox "FunctionName: ���łɋN�����Ă���Shell���擾����" & vbCrLf & _
    Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Sub

'******************************************************************************
'FunctionName:OpenIE
'Specifications�FIE���N������
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
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
'Specifications�FURL���w�肵��Web�y�[�W�ֈړ�����i���I�y�[�W�j
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub openURL()

  Dim ie As InternetExplorer

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
    
  Set ie = CreateObject("InternetExplorer.Application")

  ie.Visible = True

  '�w�肵��URL�Ɉړ�����
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
'Specifications�FIE�́uBusy�v�v���p�e�B���Ď�����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub WaitTest()

  Dim ie As InternetExplorer

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Set ie = CreateObject("InternetExplorer.Application")

  ie.Visible = True

  '�w�肵��URL�Ɉړ�����
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
'Specifications�FIE�́uReadyState�v�v���p�e�B���Ď�����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub WaitTest2()

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
   
  Dim ie As InternetExplorer

  Set ie = CreateObject("InternetExplorer.Application")

  ie.Visible = True

  '�w�肵��URL�Ɉړ�����
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
'Specifications�FWeb�y�[�W�̃^�C�g�������Ďw�肵��IE���擾����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub SearchIEI()

  Dim colSh As Object
  Dim win As Object
  Dim strTemp As String
  Dim objIE As Object
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
  
  '���݊J���Ă���IE�ƃG�N�X�v���[���[��colSh�Ɋi�[����
  Set colSh = CreateObject("Shell.Application")
  
  'colSh����Windows��1�����o��
  For Each win In colSh.Windows

    'HTMLDocument��������
    If InStr(win.document, "HTMLDocument") > 0 Then
      
      '�^�C�g���o�[��PC Watch���܂܂�邩����
      If InStr(win.document.Title, "PC Watch") > 0 Then
      
        '����objIE�Ɏ擾����win���i�[
        Set objIE = win

        '���[�v�𔲂���
        Exit For

      End If

    End If

  Next

  If objIE Is Nothing Then
  
    MsgBox "�T���Ă���IE�͂���܂���ł���"
  
  Else
  

    '�^�C�g����\������
    MsgBox objIE.document.Title & "������܂���"
    
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
'Specifications�F�l�X�ȏ����ŖړI��IE���擾����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub SearchIEI3()

  Dim colSh As Object
  Dim win As Object
  Dim strTemp As String
  Dim objIE As Object
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
  
  '���݊J���Ă���IE�ƃG�N�X�v���[���[��colSh�Ɋi�[����
  Set colSh = CreateObject("Shell.Application")
  
  'colSh����Windows��1�����o��
  For Each win In colSh.Windows
    
    strTemp = ""
    
    On Error Resume Next

    strTemp = win.document.body.innerText

    On Error GoTo 0

    '�^�C�g���o�[��PC Watch���܂܂�邩����
    If InStr(strTemp, "�A�b�v�f�[�g���") > 0 Then
    
      '����objIE�Ɏ擾����win���i�[
      Set objIE = win

      '���[�v�𔲂���
      Exit For

    End If
  
  Next

  If objIE Is Nothing Then
  
    MsgBox "�T���Ă���IE�͂���܂���ł���"
  
  Else

    '�^�C�g����\������
    MsgBox objIE.document.Title & "������܂���"
  
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
'Specifications�FIE���ǂ����̔��f�������Ƀ^�C�g���Ŕ��肷��
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub SearchIEI4()

  Dim colSh As Object
  Dim win As Object
  Dim strTemp As String
  Dim objIE As Object
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
   
  '���݊J���Ă���IE�ƃG�N�X�v���[���[��colSh�Ɋi�[����
  Set colSh = CreateObject("Shell.Application")
  
  'colSh����Windows��1�����o��
  For Each win In colSh.Windows
    
    strTemp = ""
    
    On Error Resume Next
    strTemp = win.document.Title
    On Error GoTo 0

    '�^�C�g���o�[��PC Watch���܂܂�邩����
    If InStr(strTemp, "PC Watch") > 0 Then
      
      '����objIE�Ɏ擾����win���i�[
      Set objIE = win

      '���[�v�𔲂���
      Exit For

    End If
  
  Next

  If objIE Is Nothing Then
  
    MsgBox "�T���Ă���IE�͂���܂���ł���"
  
  Else

    '�^�C�g����\������
    MsgBox objIE.document.Title & "������܂���"
    
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
'FunctionName:�őO��SetForegroundWindow
'Specifications�F�擾����IE���őO�ʂɕ\������
'Arguments�FobjIE/Object
'ReturnValue:nothing
'Note�F
'&H0:SW_HIDE �E�C���h�E���\���ɂ��A���̃E�C���h�E���A�N�e�B�u�ɂ���
'&H2:SW_MAXIMIZE�@�E�C���h�E���ő剻����
'&H3:SW_MINIMIZE�@�E�C���h�E���ŏ�������
'&H9:SW_RESTORE�@�ŏ����܂��͍ő剻����Ă����E�C���h�E�����̈ʒu�ƃT�C�Y�ɖ߂�
'******************************************************************************
Function �őO��SetForegroundWindow(objIE As Object)

  On Error GoTo Err_Trap
  
  Call m_common.MacroStart
  
  '�w�肳�ꂽ�E�C���h�E���őO�ʉ�����

  '�ŏ�������Ă���ꍇ�͌��̑傫���ɖ߂�
  If IsIconic(objIE.hWnd) Then
    
    '9��RESTORE�F�ŏ����O�̏��
    ShowWindowAsync objIE.hWnd, &H9
    
  End If
  
  'IE���őO�ʂɕ\��
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
    MsgBox "FunctionName: �őO��SetForegroundWindow" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly, "Error"

    'Clear error
    Err.Clear
  
  End If

End Function


'******************************************************************************
'FunctionName:sendKey
'Specifications�Fnothing
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
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
'Specifications�FLastPopup
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub LastPopup()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.hWnd = GetLAstActivePopup(objIE.hWnd)

    DoEvents

  Loop

 '���b�Z�[�W��\������
  MsgBox "�_�C�A���O���\�����ꂽ�B"

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
'Specifications�F���ۂɃT�C�g��̃|�b�v�A�b�v�E�C���h�E�����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub Sendkeys1()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/IE.html"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.Busy

    Sleep 100
    
    Sendkeys "{ENTER}", True

  Loop

 '���b�Z�[�W��\������
  MsgBox "Enter�L�[���������ꂽ�B"
  
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
    & Err.Description, vbOKOnly, "Sendkeys1�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Sendkeys2
'Specifications�F�t�@�C�����_�E�����[�h����BNavigate��Ƀ_�E�����[�h����t�@�C
'���𒼐ڎw�肷��Ɓu�t�@�C���̕ۑ��v�_�C�A���O���\������܂��B���̃_�C�A���O
'�����T���v���ł��B
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub Sendkeys2()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"

  '3�b�x��ł���Alt+S�𑗐M
  Sleep 3000
    
  Sendkeys "%S", True

 '���b�Z�[�W��\������
  MsgBox "�t�@�C�����ۑ����ꂽ�B"
  
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
    & Err.Description, vbOKOnly, "Sendkeys2�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Sendkeys2_2
'Specifications�F�t�@�C�����_�E�����[�h����BNavigate��Ƀ_�E�����[�h����t�@�C
'���𒼐ڎw�肷��Ɓu�t�@�C���̕ۑ��v�_�C�A���O���\������܂��B���̃_�C�A���O
'�����T���v���ł��B
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub Sendkeys2_2()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.hWnd = GetLAstActivePopup(objIE.hWnd)

    DoEvents

  Loop
    
  Sendkeys "%S", True

 '���b�Z�[�W��\������
  MsgBox "�t�@�C�����ۑ����ꂽ�B"
  
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
    & Err.Description, vbOKOnly, "Sendkeys2_2�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Checkbusy
'Specifications�F�y�[�W���J�����P�b��Ƀ|�b�v�A�b�v�E�C���h�E�Ƃ��ăt�@�C����
'�ۑ��_�C�A���O���\�������悤�ɂȂ��Ă��܂����ABusy�v���p�e�B�̕ω��̓C�~�f�B
'�G�C�g�E�C���h�E�ɂĊm�F���邱�Ƃ��o���܂��B
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub Checkbusy()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.hWnd = GetLAstActivePopup(objIE.hWnd)

    DoEvents

    '3�b�x�ށi�|�b�v�A�b�v�E�C���h�E�̕\�����Ԃɍ��킹�Ē����j
    Sleep 3000

  Loop

 '���b�Z�[�W��\������
  MsgBox "�t�@�C�����ۑ����ꂽ�Bkkk"
  
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
    & Err.Description, vbOKOnly, "Checkbusy�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Sendkeys3
'Specifications�F�t�@�C�����_�E�����[�h����BNavigate��Ƀ_�E�����[�h����t�@�C
'���𒼐ڎw�肷��Ɓu�t�@�C���̕ۑ��v�_�C�A���O���\������܂��B���̃_�C�A���O
'�����T���v���ł��B
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub Sendkeys3()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.hWnd = GetLAstActivePopup(objIE.hWnd)

    DoEvents

  Loop
    
  Sendkeys "%S", True

 '���b�Z�[�W��\������
  MsgBox "�t�@�C�����ۑ����ꂽ�Bkkk"
  
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
    & Err.Description, vbOKOnly, "Sendkeys3�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:Sendkeys3_2
'Specifications�F�t�@�C�����_�E�����[�h����BNavigate��Ƀ_�E�����[�h����t�@�C
'���𒼐ڎw�肷��Ɓu�t�@�C���̕ۑ��v�_�C�A���O���\������܂��B���̃_�C�A���O
'�����T���v���ł��B
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub Sendkeys3_2()


  On Error GoTo Err_Trap

  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")

  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"

  'busy�̊ԑҋ@
  Do While objIE.Busy

    Sleep 1

  Loop

  'busy�ƂȂ�܂őҋ@
  Do Until objIE.Busy

    Sleep 1

  Loop

  
  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.Busy

    DoEvents

  Loop

  Sendkeys "%S", True

 '���b�Z�[�W��\������
  MsgBox "�t�@�C�����ۑ����ꂽ�Bkkk"

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
    & Err.Description, vbOKOnly, "Sendkeys3_2�FError"

    'Clear error
    Err.Clear

    Call m_common.Macroend

  End If

End Sub

'******************************************************************************
'FunctionName:ShowBars
'Specifications�FIE�̕\���𐧌䂷��
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub ShowBars()

  On Error GoTo Err_Trap

  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")

  '����
  objIE.Visible = True
  
  '
  objIE.Toolbar = True
  
  '
  objIE.AddressBar = True
  
  '
  objIE.MenuBar = True
    
  '
  objIE.StatusBar = True
  
 '���b�Z�[�W��\������
  MsgBox "�e��o�[���\�����ꂽ�B"

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
    & Err.Description, vbOKOnly, "ShowBars�FError"

    'Clear error
    Err.Clear

    Call m_common.Macroend

  End If

End Sub

'******************************************************************************
'FunctionName:ChangeSizeAndLocation
'Specifications�F�E�C���h�E�̃T�C�Y�ƈʒu���w�肷��
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub ChangeSizeAndLocation()

  On Error GoTo Err_Trap

  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")

  '����
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
 
 '���b�Z�[�W��\������
  MsgBox "�E�C���h�E�̃T�C�Y�ƈʒu���w�肳�ꂽ�B"

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
    & Err.Description, vbOKOnly, "ChangeSizeAndLocation�FError"

    'Clear error
    Err.Clear

    Call m_common.Macroend

  End If

End Sub

'******************************************************************************
'FunctionName:ExecInvisible
'Specifications�F�E�C���h�E��\���������������s
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub ExecInvisible()

  On Error GoTo Err_Trap

  Call m_common.MacroStart

  Dim objIE As Object

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")

  '�����I�t
  objIE.Visible = False
  
  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://yahoo.co.jp/"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

 
 '���b�Z�[�W��\������
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
    & Err.Description, vbOKOnly, "ExecInvisible�FError"

    'Clear error
    Err.Clear

    Call m_common.Macroend

  End If

End Sub

'******************************************************************************
'FunctionName:useAnchor
'Specifications�FIE��ʏ�̃n�C�p�[�����N���g����Web�y�[�W���ړ�����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub useAnchor()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim anchor As HTMLAnchorElement

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-7.html"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

  '�����N�̐ݒ肳�ꂽ�����񂩂珈���Ώۂ���������
  For Each anchor In objIE.document.getElementsByTagName("A")

    If anchor.innerText = "�₫���΃p�� vs �g���p��" Then
      
      '�n�C�p�[�����N���N���b�N
      anchor.Click

      '���b�Z�[�W��\������
      MsgBox "�₫���΃p�� vs �g���p���y�[�W���\�����ꂽ�B"
      
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
    & Err.Description, vbOKOnly, "useAnchor�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:useButton
'Specifications�F�{�^���𑀍삷��
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub useButton()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim button As HTMLButtonElement

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-8.html"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

  '�{�^���\�ʂ̕����񂩂珈���Ώۂ���������
  For Each button In objIE.document.getElementsByTagName("INPUT")
    
    If button.Type = "button" And button.Value = "�{�^���Q" Then
      
      '�{�^�����N���b�N
      button.Click

     '���b�Z�[�W��\������
      MsgBox "�{�^�����N���b�N���ꂽ�B"
      
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
    & Err.Description, vbOKOnly, "useButton�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

''******************************************************************************
''FunctionName:openURL
''Specifications�FIE��ʏ�̉摜��؂�ւ���B
''Arguments�Fnothing
''ReturnValue:nothing
''Note�F
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
'  'IE���J���ăt�@�C���̕ۑ�URL���J��
'  Set objIE = CreateObject("InternetExplorer.Application")
'  '����
'  objIE.Visible = True
'
'  '�w�肵��URL�Ɉړ�����
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
'          '0.2�b��~��A��ʂ����ɖ߂��A�ēx0.2�b��~
'          Sleep 200
'
'          .src = Src1
'
'          '0.2�b��~��A��ʂ����ɖ߂��A�ēx0.2�b��~
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
'    & Err.Description, vbOKOnly, "openURL�FError"
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
'Specifications�F�^�O�̈ꗗ�\���쐬����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub GetTable3()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim button As HTMLButtonElement

  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-10_3.html"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
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
    & Err.Description, vbOKOnly, "GetTable3�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:MakeList
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub MakeList(objIE As InternetExplorer)
  
  Dim n As Long '�^�O�̒ʂ��ԍ�
  Dim r As Long 'td,th�^�O�̒ʂ��ԍ�
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  '�^�O�i�[�p
  Dim wslist As Worksheet

  On Error GoTo Err_Trap
  
  Debug.Print "(Function: MakeList START)"
  '�ϐ�������
  n = 0
  r = 0
  
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/�W��"
    
    Set Doc = objIE.document
    
  End With
    Debug.Print Doc.all.Length - 1
  '�{�^���\�ʂ̕����񂩂珈���Ώۂ���������
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
        
        '����
        If r = 1 Then
        
          'number
          wslist.Cells(r, 1) = "Number"
          
          '�^�O�̖��O
          wslist.Cells(r, 2) = "�^�O�̖��O"
        
          '�^�O�̒ʂ��ԍ�
          wslist.Cells(r, 3) = "�^�O�̒ʂ��ԍ�"
        
          'td,th�^�O�̒ʂ��ԍ�
          wslist.Cells(r, 4) = "td,th�^�O�̒ʂ��ԍ�"
        
          '�e�L�X�g(SOURCEINDEX)
          wslist.Cells(r, 5) = "SOURCEINDEX"
          
          '�e�L�X�g(innertext)
          wslist.Cells(r, 6) = "�e�L�X�g(innertext)"
          
          '�e�L�X�g(outertext)
          wslist.Cells(r, 7) = "�e�L�X�g(outertext)"
      
          'HTML(outerhtml)
          wslist.Cells(r, 8) = "HTML(outerhtml)"
      
        '�擾���
        ElseIf r > 1 Then
          'number
          wslist.Cells(r, 1) = r - 1
          
          '�^�O�̖��O
          wslist.Cells(r, 2) = .tagName
        
          '�^�O�̒ʂ��ԍ�
          wslist.Cells(r, 3) = n
        
          'td,th�^�O�̒ʂ��ԍ�
          wslist.Cells(r, 4) = r
        
          'SOURCEINDEX
          wslist.Cells(r, 5) = .sourceIndex
        
          '�e�L�X�g(innertext)
          wslist.Cells(r, 6) = .innerText
          
          '�e�L�X�g(outertext)
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
    & Err.Description, vbOKOnly, "MakeList2�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:MakeList2
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F537�΂���855�Ԗڂ܂ł̃^�O�𒲐�����
'td�^�O�Ȃ�擾���A1�񂸂E�ɂ��炵�Ȃ���Z���ɏ�������
'16�擾�����玟�̍s��1��ڂɈڂ�B
'******************************************************************************
Sub MakeList2(objIE As InternetExplorer)
  
  Dim n As Long '�^�O�̒ʂ��ԍ�
  Dim r As Long 'td,th�^�O�̒ʂ��ԍ�
  Dim i As Long '
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  '�^�O�i�[�p
  Dim wslist As Worksheet

  On Error GoTo Err_Trap
  
  '�ϐ�������
  n = 0
  r = 0
  i = 0
  
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/�W��"
    
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
    & Err.Description, vbOKOnly, "MakeList2�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:MakeList3
'Specifications�F�e�[�u�����ς̏ꍇ�̎擾���@
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub MakeList3(objIE As InternetExplorer)
  
  Dim n As Long '�^�O�̒ʂ��ԍ�
  Dim r As Long 'td,th�^�O�̒ʂ��ԍ�
  Dim i As Long '
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  '�^�O�i�[�p
  Dim wslist As Worksheet
  Dim StartTag As Long
  Dim FinishTag As Long
  
  On Error GoTo Err_Trap
  
  '�ϐ�������
  n = 0
  r = 0
  i = 0
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/�W��"
    
    Set Doc = objIE.document
    
  End With
  
  '�h�L�������g�\���^�O���P������
  For i = 0 To Doc.all.Length - 1
    
    'th�^�O�Ȃ�
    If Doc.all(i).tagName = "TH" Then
    
      If Doc.all(i).innerText = "�t��" Then
      
        StartTag = 1
        
        Exit For
        
      End If
      
    End If
      
  Next i
    
  '�h�L�������g�\���^�O���P������
  For i = StartTag To Doc.all.Length - 1
    
    'th�^�O�Ȃ�
    If Doc.all(i).tagName = "TH" Then
    
      If Doc.all(i).innerText = "���[�J�[" Then
        
        FinishTag = 1

        Exit For
        
      End If
      
    End If
      
  Next i
  
  '�h�L�������g�\���^�O���P������
  For i = StartTag To FinishTag
    
    'td�^�O�Ȃ�
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
    & Err.Description, vbOKOnly, "MakeList3�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub


'******************************************************************************
'FunctionName:MakeList4
'Specifications�F��荂���ɏ�������
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub MakeList4(objIE As InternetExplorer)
  
  Dim n As Long '�^�O�̒ʂ��ԍ�
  Dim r As Long 'td,th�^�O�̒ʂ��ԍ�
  Dim i As Long '
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  '�^�O�i�[�p
  Dim ObjTD As Object  '�^�O�i�[�p
  Dim wslist As Worksheet

  On Error GoTo Err_Trap
  
  '�ϐ�������
  n = 0
  r = 0
  i = 0
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/�W��"
    
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
    & Err.Description, vbOKOnly, "MakeList4�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub


'******************************************************************************
'FunctionName:MakeList5
'Specifications�F���ׂẴe�[�u���̈ꗗ�\���쐬����R�[�h�����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub MakeList5(objIE As InternetExplorer)
  
  Dim n As Long '�^�O�̒ʂ��ԍ�
  Dim r As Long 'td,th�^�O�̒ʂ��ԍ�
  Dim c As Long '�J����
  Dim i As Long '
  Dim Doc As HTMLDocument
  Dim ObjTag As Object  '�^�O�i�[�p
  Dim ObjTD As Object  '�^�O�i�[�p
  Dim wslist As Worksheet

  On Error GoTo Err_Trap
  
  '�ϐ�������
  n = 0
  r = 0
  i = 0
  Set wslist = ThisWorkbook.Worksheets("list")

  With wslist
    
    .Cells.ClearContents
    .Cells.NumberFormatLocal = "G/�W��"
    
   Set Doc = objIE.document
    
  End With
  
  For i = 0 To Doc.all.Length - 1

    'td�^�O��th�^�O
    If Doc.all(i).tagName = "TH" Or Doc.all(i).tagName = "TD" Then
  
      wslist.Cells(r, c) = Doc.all(i).innerText

      c = c + 1
    'tr�^�O
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
    & Err.Description, vbOKOnly, "MakeList5�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:useScript
'Specifications�F�X�N���v�g�����s����BWeb�y�[�W�Ƀ��b�Z�[�W�{�b�N�X��\������
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub useScript()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim button As HTMLButtonElement
  Dim pwin As HTMLWindow2
  
  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-13.html"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

  Set pwin = objIE.document.parentWindow
  
  pwin.alert ("VBA����alert�����s")
  
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
    & Err.Description, vbOKOnly, "useScript�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName:useScript2
'Specifications�F�X�N���v�g�����̊�����҂�����VBA�̌㑱���������s����
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub useScript2()
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim objIE As Object
  Dim button As HTMLButtonElement
  Dim pwin As HTMLWindow2
  
  'IE���J���ăt�@�C���̕ۑ�URL���J��
  Set objIE = CreateObject("InternetExplorer.Application")
  
  '����
  objIE.Visible = True

  '�w�肵��URL�Ɉړ�����
  objIE.navigate "http://book.impress.co.jp/appended/3384/4-13.html"

  '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
  Do While objIE.Busy Or objIE.readyState < READYSTATE_COMPLETE

    Debug.Print objIE.Busy & ":" & objIE.readyState

    DoEvents

  Loop

  Set pwin = objIE.document.parentWindow
  
  pwin.execScript "showMessage('VBA����showMessage�����s')"
  
  '**********************************************************************
  'OK�{�^�����N���b�N����IE�̃��b�Z�[�W�����O��4�����s�����悤��
  '����ɂ͏�L�̏������ȉ��̗l�ɏ�������
  'pwin.setTimeout "showMessage('VBA����showMessage��񓯊����s')", 0
  '**********************************************************************
  
  MsgBox "VBA�̌㑱����"
  
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
    & Err.Description, vbOKOnly, "useScript2�FError"

    'Clear error
    Err.Clear
  
    Call m_common.Macroend
  
  End If

End Sub

'******************************************************************************
'FunctionName: getElementList
'Specifications�F'HTML�\�[�X�S�̂��擾����
'Arguments�FgetHTMLString / String
'ReturnValue:nothing
'Note�F
'******************************************************************************
Public Function getHTMLString(ByVal objIE As InternetExplorer) As String
  
  Dim htdoc As HTMLDocument
  Dim ret As String
  Dim elle As IHTMLElement
  
  Set htdoc = objIE.document
  
  'HTML�\�[�X�S�̂��擾����
  Set elle = htdoc.getElementByTagName("HTML")(0)
  
'  ret = htdoc.getElementbyTAgName("HTML")(0).outerHTML & vbCrLf
  ret = elle.outerHTML & vbCrLf

  Set htdoc = Nothing


  getHTMLString = ret

End Function


'******************************************************************************
'FunctionName: getElementList
'Specifications�F
'Arguments�Fhtdoc / HTMLDocument
'ReturnValue:nothing
'Note�F
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
'Specifications�F
'Arguments�Fhtdoc / HTMLDocument
'ReturnValue:nothing
'Note�F
'******************************************************************************
Public Sub printMain01(ByVal objIE As InternetExplorer)

  On Error GoTo Err_Trap

  Dim HTMLstring  As String
  Dim FileName    As String
  Dim FileNum     As Long

  'HTML�S���擾
'  HTMLstring = getHTMLString(objIE)
  HTMLstring = getHTMLString(objIE)
  
  FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & ".txt"
  
  '�t�@�C���ԍ��̎擾
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
'Specifications�F
'Arguments�Fhtdoc / HTMLDocument
'ReturnValue:nothing
'Note�F
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
  
  '���^�[���l����؂���ƊK�w���ŏ�����
  ret = "-------------------------------------------------------------" & vbCrLf
  
  ret = ret & "[" & Trim(str(depth)) & "�K�w]" & vbCrLf
  
  Dim i As Integer
  
  'If HTML can be retrieved
  If Not htdoc Is Nothing Then
  
    'Frame and document information (�t���[���ƕ����̏��)
'    ret = ret & htdoc.Title & " | " & htdoc.Location & " (" & container.Name & ")" & vbCrLf
    ret = ret & htdoc.Title & " | " & htdoc.Location & vbCrLf
    
    ret = ret & "-------------------------------------------------------------" & vbCrLf
  
'    'Obtain a list of screen components
'    ret = ret & htdoc.getElementList(htdoc) & vbCrLf


'    ret = ret & "-------------------------------------------------------------" & vbCrLf
    
    '(HTML�^�O�v�f(��ʂɈ��))
    ret = ret & objIE.document.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
    
  
    For i = 0 To objIE.document.frames.Length - 1
    
      'Recurses if frames are present (�t���[��������ꍇ�͍ċA����)
      ret = ret & htdoc.getHTMLString(htdoc.frames(i), depth + 1)
    
    Next i
  
  Else
  
    ret = ret & "-------------------------------------------------------------" & vbCrLf
    
    'Output error information if HTML could not be obtained
    '(HTML���擾�o���Ȃ������ꍇ�̓G���[�����o��)
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
'Specifications�F�t���[�����܂߂����ׂĂ�HTML���o�͂���
'Arguments�Fhtdoc / HTMLDocument
'ReturnValue:nothing
'Note�F
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

  '���^�[���l����؂���ƊK�w���ŏ�����
  ret = "-------------------------------------------------------------" & vbCrLf
  
  '(HTML�^�O�v�f(��ʂɈ��))
  ret = ret & objIE.document.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
  
  
  For i = 0 To objIE.document.frames.Length - 1
    
    'Recurses if frames are present (�t���[��������ꍇ�͍ċA����)
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
'Specifications�F�G���[�Œ��������f����Ȃ��悤�ɂ��A���ׂĂ�HTML���o�͂���
'Arguments�Fhtdoc / HTMLDocument
'ReturnValue:nothing
'Note�F
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

  '���^�[���l����؂���ƊK�w���ŏ�����
  ret = "-------------------------------------------------------------" & vbCrLf
  
  ret = ret & objIE.document.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
  
  If Not htdoc Is Nothing Then
    
    '(HTML�^�O�v�f(��ʂɈ��))
    ret = ret & objIE.document.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
  
    For i = 0 To objIE.document.frames.Length - 1
      
      'Recurses if frames are present (�t���[��������ꍇ�͍ċA����)
      ret = ret & getHTMLString3(objIE.document.frames(i), objIE)
      
    Next i
  
  Else
  
    'HTML���擾�o���Ȃ������ꍇ�̓G���[�����o�͂���
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
'�yobjIE.document.getElementsByTagName("h2")(0).outerHTML�z
'objIE = InternetExplorer�I�u�W�F�N�g
'document = HTML�h�L�������g�̃I�u�W�F�N�g(Document�I�u�W�F�N�g)
'getElementsByTagName("h2") = HTML�h�L�������g���̂��ׂĂ�h2�v�f(GetElementsByTagName���\�b�h)
'getElementsByTagName("h2")(0) = h2�v�f�R���N�V������1�Ԗڂ�h2�v�f�I�u�W�F�N�g
'outerHTML = 1�Ԗڂ�h2�v�f�I�u�W�F�N�g�̗v�f�^�O�Ƃ��̒��Ɋ܂܂��HTML�R�[�h


'******************************************************************************
'FunctionName: getElementList2
'Specifications�F
'Arguments�Fhtdoc / HTMLDocument
'ReturnValue:nothing
'Note�F
'******************************************************************************
Public Function getElementList2(ByVal htdoc As HTMLDocument)
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim ret As String
  Dim element As Object
  Dim i  As Long    '���i�ԍ��p�ϐ���錾����
  
  '�w�b�_�[�̐擪�ɔԍ����������ڂ�ǉ�����
  ret = "#" & vbTab & "�^�O" & vbTab & "Type" & vbTab & "ID" & vbTab & "Name" & _
    vbTab & "Value" & vbCrLf

  '�h�L�������g�S�v�f�ɑ΂��ď���
  For Each element In htdoc.all
  
    Select Case UCase(element.tagName)
      
      'Evaluate whether the tag type is INPUT/TEXTAREA/SELECT.
      Case "INPUT", "TEXTAREA", "SELECT"
        
        '��ʂԂЂ񂶂傤�ق��̐擪�ɔԍ����L�^����
        ret = ret & CStr(i) & vbTab & element.tagName & vbTab & element.Type & vbTab & _
        element.ID & vbTab & element.Name & vbTab & element.Value & vbCrLf

        '��ʂɔԍ��������߂�
        If UCase(element.Type) <> "HIDDEN" Then
        
          element.outerHTML = element.outerHTML & "&nbsp;<bstyle=""color:blue;"">[" & CStr(i) & "]</b>"
        
        End If
        
        '�ԍ����C���N�������g����
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
'Specifications�F
'Arguments�Fhtdoc / HTMLDocument
'ReturnValue:nothing
'Note�F
'******************************************************************************
Public Function getElementList3(ByVal htdoc As HTMLDocument)
  
  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  Dim ret As String
  Dim element As Object
  Dim i  As Long    '���i�ԍ��p�ϐ���錾����
  
  '�w�b�_�[�̐擪�ɔԍ����������ڂ�ǉ�����
  ret = "#" & vbTab & "�^�O" & vbTab & "Type" & vbTab & "ID" & vbTab & "Name" & _
    vbTab & "Value" & vbCrLf

  '�h�L�������g�S�v�f�ɑ΂��ď���
  For Each element In htdoc.all
  
    Select Case UCase(element.tagName)
      
      'Evaluate whether the tag type is INPUT/TEXTAREA/SELECT.
      Case "INPUT", "TEXTAREA", "SELECT", "A", "DIV", "SCRIPT", "TD", "P", "TR", _
        "SPAN", "STRONG", "BR", "TABLE", "TBODY", "IMG", "OPTION", "CENTER", "HEAD", _
        "BODY", "LABEL", "LI", "UR", "fieldset", "form", "H1", "H2", "H3", "H4", "H5", _
        "IFRAME", "THEAD", "BODY", "LEFT", "RIGHT", "HTML"
      
        '��ʂԂЂ񂶂傤�ق��̐擪�ɔԍ����L�^����
        ret = ret & CStr(i) & vbTab & element.tagName & vbTab & element.Type & vbTab & _
        element.ID & vbTab & element.Name & vbTab & element.Value & vbCrLf

        '��ʂɔԍ��������߂�
        If UCase(element.Type) <> "HIDDEN" Then
        
          element.outerHTML = element.outerHTML & "&nbsp;<bstyle=""color:blue;"">[" & CStr(i) & "]</b>"
        
        Else
          
          element.outerHTML = element.outerHTML & "&nbsp;<bstyle=""color:blue;"">[" & CStr(i) & "]</b>"
        
        End If
        
        '�ԍ����C���N�������g����
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



