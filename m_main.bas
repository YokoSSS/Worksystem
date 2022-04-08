Attribute VB_Name = "m_main"
Option Explicit
'//Sleep�@�\���g��API
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)


'******************************************************************************
'FunctionName:
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
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
'Specifications�F
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub Login()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IE�I�u�W�F�N�g������
Dim htmlDoc    As HTMLDocument     'HTML�h�L�������g�I�u�W�F�N�g������
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElement�I�u�W�F�N�g������
Dim eltext     As IHTMLElement '�g�p���Ă��Ȃ�
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'�f�X�N�g�b�v��"WORK"�t�H���_���쐬����
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****���O�C��*****
  '�f�X�N�g�b�v�Ƀt�H���_�p�X���E��
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  '�G�C�W�X���T�[�`�l��
  IDusername = "1124_senoo"   '���O�C�����[�U�[�l�[��
  IDpassword = "to4lklp7"     '���O�C���p�X���[�h

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IE�I�u�W�F�N�g���Z�b�g����
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IE��\��
    objIE.Visible = True

    'IE��URL���J��
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '�ǂݍ��ݑ҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    Set htmlDoc = objIE.document
    'ID���Z�b�g
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
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
    
    '�t�@�C���ԍ��̎擾
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
'Specifications�F�t���[�����܂߂����ׂĂ�HTML���o�͂���
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub SampleLogin002()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IE�I�u�W�F�N�g������
Dim htmlDoc    As HTMLDocument     'HTML�h�L�������g�I�u�W�F�N�g������
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElement�I�u�W�F�N�g������
Dim eltext     As IHTMLElement '�g�p���Ă��Ȃ�
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'�f�X�N�g�b�v��"WORK"�t�H���_���쐬����
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****���O�C��*****
  '�f�X�N�g�b�v�Ƀt�H���_�p�X���E��
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  '�G�C�W�X���T�[�`�l��
  IDusername = "1124_senoo"   '���O�C�����[�U�[�l�[��
  IDpassword = "to4lklp7"     '���O�C���p�X���[�h

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IE�I�u�W�F�N�g���Z�b�g����
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IE��\��
    objIE.Visible = True

    'IE��URL���J��
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '�ǂݍ��ݑ҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    Set htmlDoc = objIE.document
    'ID���Z�b�g
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
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
    
    '�t�@�C���ԍ��̎擾
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
'Specifications�F�G���[�Œ��������f����Ȃ��悤�ɂ��A���ׂĂ�HTML���o�͂���
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub getHTMLString3Login003()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IE�I�u�W�F�N�g������
Dim htmlDoc    As HTMLDocument     'HTML�h�L�������g�I�u�W�F�N�g������
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElement�I�u�W�F�N�g������
Dim eltext     As IHTMLElement '�g�p���Ă��Ȃ�
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'�f�X�N�g�b�v��"WORK"�t�H���_���쐬����
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****���O�C��*****
  '�f�X�N�g�b�v�Ƀt�H���_�p�X���E��
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  '�G�C�W�X���T�[�`�l��
  IDusername = "1124_senoo"   '���O�C�����[�U�[�l�[��
  IDpassword = "to4lklp7"     '���O�C���p�X���[�h

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IE�I�u�W�F�N�g���Z�b�g����
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IE��\��
    objIE.Visible = True

    'IE��URL���J��
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '�ǂݍ��ݑ҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    Set htmlDoc = objIE.document
    'ID���Z�b�g
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
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
    
    '�t�@�C���ԍ��̎擾
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
'Specifications�F��ʂƉ�͏����ƍ����邽�߂̔ԍ���\������
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub getHTMLStringLogin004()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IE�I�u�W�F�N�g������
Dim htmlDoc    As HTMLDocument     'HTML�h�L�������g�I�u�W�F�N�g������
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElement�I�u�W�F�N�g������
Dim eltext     As IHTMLElement '�g�p���Ă��Ȃ�
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'�f�X�N�g�b�v��"WORK"�t�H���_���쐬����
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****���O�C��*****
  '�f�X�N�g�b�v�Ƀt�H���_�p�X���E��
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  '�G�C�W�X���T�[�`�l��
  IDusername = "1124_senoo"   '���O�C�����[�U�[�l�[��
  IDpassword = "to4lklp7"     '���O�C���p�X���[�h

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IE�I�u�W�F�N�g���Z�b�g����
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IE��\��
    objIE.Visible = True

    'IE��URL���J��
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '�ǂݍ��ݑ҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    Set htmlDoc = objIE.document
    'ID���Z�b�g
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    
    i = 0
    '
    str = getElementList2(htmlDoc)
  
  
    Debug.Print str
    
    FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & "_" & j - 1 & ".txt"
    
    '�t�@�C���ԍ��̎擾
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
'Specifications�F��ʂƉ�͏����ƍ����邽�߂̔ԍ���\������
'Arguments�Fnothing
'ReturnValue:nothing
'Note�F
'******************************************************************************
Sub getHTMLStringLogin005()

Dim IDusername As String
Dim IDpassword As String
Dim objIE      As InternetExplorer 'IE�I�u�W�F�N�g������
Dim htmlDoc    As HTMLDocument     'HTML�h�L�������g�I�u�W�F�N�g������
Dim elFormID   As IHTMLElement, elFormpass As IHTMLElement 'IHTMLElement�I�u�W�F�N�g������
Dim eltext     As IHTMLElement '�g�p���Ă��Ȃ�
Dim elbutton   As HTMLFormElement
Dim wb         As Workbook
Dim ReviewID   As String
Dim i          As Long
Dim j          As Long
Dim lastrow    As Long

Dim wsh        As Variant
Dim Path       As String
'�f�X�N�g�b�v��"WORK"�t�H���_���쐬����
Const fdn As String = "WORK"

Dim str As String
Dim FileNum     As Long
Dim FileName    As String


  On Error GoTo Err_Trap
  
  Call m_common.MacroStart

  '*****���O�C��*****
  '�f�X�N�g�b�v�Ƀt�H���_�p�X���E��
  Set wsh = CreateObject("WScript.Shell")
  Path = wsh.SpecialFolders("Desktop") & "\" & fdn & "\"

  Debug.Print "(" & Path & ")"
  '�G�C�W�X���T�[�`�l��
  IDusername = "1124_senoo"   '���O�C�����[�U�[�l�[��
  IDpassword = "to4lklp7"     '���O�C���p�X���[�h

  lastrow = ThisWorkbook.Worksheets("list2").Cells(Rows.Count, 1).End(xlUp).Row
  
  For j = 2 To lastrow
  
    ReviewID = ThisWorkbook.Worksheets("list2").Cells(j, 1).Value
    
    Debug.Print "ReviewID: " & ReviewID
    
    'IE�I�u�W�F�N�g���Z�b�g����
    Set objIE = CreateObject("Internetexplorer.Application")
  
    'IE��\��
    objIE.Visible = True

    'IE��URL���J��
    objIE.navigate "https://csc.ajis-group.co.jp/jp/login.php"
    
    '�ǂݍ��ݑ҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
  
    Sleep 10000
    
    'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    Set htmlDoc = objIE.document
    'ID���Z�b�g
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    objIE.navigate "https://csc.ajis-group.co.jp/edit-entire-crit.php?CritID=" & ReviewID
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    'objIE
    Set elFormID = htmlDoc.getElementById("username")
    Set elFormpass = htmlDoc.getElementById("password")
    Set elbutton = objIE.document.getElementById("do_login")
    
    'ID�����
    elFormID.Value = IDusername
    'Pass�����
    elFormpass.Value = IDpassword
    '���M
    elbutton.Click
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
    Sleep 15000
    
    
    i = 0
    
    Call MakeList(objIE)
    
    '
    str = getElementList3(htmlDoc)
  
  
    Debug.Print str
    
    FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & "_" & j - 1 & ".txt"
    
    '�t�@�C���ԍ��̎擾
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

