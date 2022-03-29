Attribute VB_Name = "m_basfileExport"
'Option Explicit
'
''******************************************************************************
''FunctionName:�P�̃t�@�C����Bas�t�@�C�����G�N�X�|�[�g����
''Specifications�F�P�̃t�@�C����Bas�t�@�C�����G�N�X�|�[�g����
''Arguments�Fnothing
''ReturnValue:nothing
''Note�F
''******************************************************************************
'Sub ��̃t�@�C����Bas�t�@�C�����G�N�X�|�[�g����()
'
'Dim module                  As VBComponent      '// ���W���[��
'Dim moduleList              As VBComponents     '// VBA�v���W�F�N�g�̑S���W���[��
'Dim extension                                   '// ���W���[���̊g���q
'Dim sPath                   As String           '// �����Ώۃu�b�N�̃p�X
'Dim sFilePath               As String           '// �G�N�X�|�[�g�t�@�C���p�X
'Dim sFoldPath               As String           '// �����Ώۃu�b�N�̃t�H���_�p�X
'Dim TargetBook              As Workbook         '// �����Ώۃu�b�N�I�u�W�F�N�g
'
'  On Error GoTo Err_Trap
'
'  Call m_common.�}�N���J�n
'
'  '�����Ώۃu�b�N�̑I��
'  If GET�t�@�C��(sPath) = False Then Exit Sub
'
'  '�����Ώۃu�b�N�̃t�H���_����
'  sFoldPath = Mid(sPath, 1, Len(sPath) - 5)
'
'  If �w�肵���p�X�Ńt�H���_����(sFoldPath) = False Then
'
'    Call m_common.�}�N���I��
'
'    Exit Sub
'
'  End If
'
'  ThisWorkbook.Worksheets("main").Range("main_file1") = sFoldPath
'
'  '�����Ώۃu�b�N�̃I�u�W�F�N�g����
'  Workbooks.Open stritems:=sPath
'
'  Set TargetBook = ActiveWorkbook
'
'  '�����Ώۃu�b�N�̃��W���[���ꗗ���擾
'  Set moduleList = TargetBook.VBProject.VBComponents
'
'  '// VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
'  For Each module In moduleList
'    '// �N���X
'    If (module.Type = vbext_ct_ClassModule) Then
'        extension = "cls"
'    '// �t�H�[��
'    ElseIf (module.Type = vbext_ct_MSForm) Then
'        '// .frx���ꏏ�ɃG�N�X�|�[�g�����
'        extension = "frm"
'    '// �W�����W���[��
'    ElseIf (module.Type = vbext_ct_StdModule) Then
'        extension = "bas"
'    '// ���̑�
'    Else
'        '// �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
'        GoTo CONTINUE
'    End If
'
'    '// �G�N�X�|�[�g���{
'    sFilePath = sFoldPath & "\" & module.Name & "." & extension
'
'    Call module.Export(sFilePath)
'
'    '// �o�͐�m�F�p���O�o��
'    Debug.Print sFilePath
'
'CONTINUE:
'    Next module
'
'  '�u�b�N�����
'  TargetBook.Close SaveChanges:=False
'
'  '���������
'  Set TargetBook = Nothing
'
'  MsgBox "Done!!�@�������I�����܂��B", vbInformation, "�����I��"
'
'  Call m_common.�}�N���I��
'
'  Exit Sub
'
'Err_Trap:
'
'  'When an error occurs, display the contents of the error in a message box.
'  If Err.Number <> 0 Then
'    '
'    Debug.Print Err.Number & " " & Err.Description
'    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"
'
'    'Clear error
'    Err.Clear
'
'  End If
'
'  Call m_common.�}�N���I��
'
'End Sub
'
''******************************************************************************
''FunctionName:���t�H���_�����t�@�C����Bas�t�@�C�����G�N�X�|�[�g����
''Specifications�F�����t�H���_�ɂ��镡���̃}�N���t�@�C����Bas�t�@�C����
''                �G�N�X�|�[�g����
''Arguments�Fnothing
''ReturnValue:nothing
''Note�F
''******************************************************************************
'Sub ���t�H���_�����t�@�C����Bas�t�@�C�����G�N�X�|�[�g����()
'
'Dim module                  As VBComponent      '// ���W���[��
'Dim moduleList              As VBComponents     '// VBA�v���W�F�N�g�̑S���W���[��
'Dim extension                                   '// ���W���[���̊g���q
'Dim sPath                   As String           '// �����Ώۃu�b�N�̃p�X
'Dim sFilePath               As String           '// �G�N�X�|�[�g�t�@�C���p�X
'Dim sFoldPath               As String           '// �����Ώۃu�b�N�̃t�H���_�p�X
'Dim sFdPath                 As String           '// �����Ώۃu�b�N�̃t�H���_�p�X
'Dim TargetBook              As Workbook         '// �����Ώۃu�b�N�I�u�W�F�N�g
'Dim FSO As New FileSystemObject
'Dim objFiles As File
'Dim objFolders As Folder
'Dim cntlist As Long
'
'  On Error GoTo Err_Trap
'
'  Call m_common.�}�N���J�n
'
'  cntlist = 1
'
'  '�ꊇ�����Ώۃt�H���_�̑I��
'  If GET�t�H���_(sFdPath) = False Then Exit Sub
'
'  '�w��t�H���_��bas�t�H���_���Ȃ�������쐬���A�������珈���������I������
'  If FolderExists(sFdPath & "\" & "bas") = True Then
'
'    MsgBox "�w�肵���t�H���_��bas�t�H���_������܂����B" & vbCrLf & _
'        "�������I�����܂��B", vbCritical, "�����I��"
'
'    Call m_common.�}�N���I��
'
'    Exit Sub
'
'  ElseIf FolderExists(sFdPath & "\" & "bas") = False Then
'
'    'bas�t�H���_���Ȃ�������쐬
'    If �w�肵���p�X�Ńt�H���_����(sFdPath & "\" & "bas") = False Then Exit Sub
'
'  End If
'
'  '�t�@�C�����̎擾
'  For Each objFiles In FSO.GetFolder(sFdPath).Files
'
'    Debug.Print "objFiles: " & objFiles
'
'    '�}�N���t�@�C����Ώۂɏ���������
'    If FSO.GetExtensionName(objFiles.Name) = "xlsm" And _
'      objFiles.Name <> ThisWorkbook.Name And _
'       Not (objFiles.Name Like "*~$*") Then
'
'      'list�V�[�g�Ƀt�@�C�������o�͂��邽�߂̍s�J�E���g
'      cntlist = cntlist + 1
'
'      '���X�g�Ɉꊇ�����Ώۃt�H���_���t�@�C���������o��
'      'If setFileList(searchPath, cntlist) = False Then Exit For
'
'      '�����Ώۃu�b�N�̃t�H���_����
'      sFoldPath = sFdPath & "\" & "bas" & "\" & FSO.GetBaseName(objFiles.Name)
'
'      Debug.Print "�t�H���_: " & sFoldPath
'
'      If �w�肵���p�X�Ńt�H���_����(sFoldPath) = False Then Exit Sub
'        'list�V�[�g�Ƀt�@�C�������o�͂���
'
'        With ThisWorkbook.Worksheets("list")
'
'          'idn
'          .Cells(cntlist, Range("list_idn").Column) = cntlist - 1
'          'stritems
'          .Cells(cntlist, Range("list_stritems").Column) = objFiles.Name
'          'filepath
'          .Cells(cntlist, Range("list_filepath").Column) = list_failepath
'          '�i�[bas�t�H���_path
'          .Cells(cntlist, Range("list_baspath").Column) = list_failepath
'
'        End With
'
'        '�����Ώۃu�b�N�̃I�u�W�F�N�g����
'        Workbooks.Open stritems:=objFiles
'
'        Set TargetBook = ActiveWorkbook
'
'        '�����Ώۃu�b�N�̃��W���[���ꗗ���擾
'        Set moduleList = TargetBook.VBProject.VBComponents
'
'        'VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
'        For Each module In moduleList
'          '�N���X
'          If (module.Type = vbext_ct_ClassModule) Then
'              extension = "cls"
'          '�t�H�[��
'          ElseIf (module.Type = vbext_ct_MSForm) Then
'              '.frx���ꏏ�ɃG�N�X�|�[�g�����
'              extension = "frm"
'          '�W�����W���[��
'          ElseIf (module.Type = vbext_ct_StdModule) Then
'              extension = "bas"
'          '���̑�
'          Else
'              '�G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
'
'              GoTo CONTINUE
'
'          End If
'
'          '�G�N�X�|�[�g���{
'          sFilePath = sFoldPath & "\" & module.Name & "." & extension
'
'        Call module.Export(sFilePath)
'
'        '�o�͐�m�F�p���O�o��
'        Debug.Print sFilePath
'
'CONTINUE:
'        Next module
'
'        '�u�b�N�����
'        TargetBook.Close SaveChanges:=False
'
'        '�u�b�N��bas�ֈړ�������
'        Name objFiles As sFdPath & "\" & "bas" & "\" & objFiles.Name
'
'        '���������
'        Set TargetBook = Nothing
'
'    End If
'
'
'  Next objFiles
'
'  MsgBox "Done!!�@�������I�����܂��B", vbInformation, "�����I��"
'
'
'  Call m_common.�}�N���I��
'
'  Exit Sub
'
'Err_Trap:
'
'  'When an error occurs, display the contents of the error in a message box.
'  If Err.Number <> 0 Then
'    '
'    Debug.Print Err.Number & " " & Err.Description
'    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"
'
'    'Clear error
'    Err.Clear
'
'  End If
'
'  Call m_common.�}�N���I��
'
'End Sub
'
''******************************************************************************
''FunctionName:���t�H���_�̃T�u�t�H���_�܂ޕ����̃}�N���t�@�C����Bas�t�@�C��
''             ���G�N�X�|�[�g����
''Specifications�F�����t�H���_�ɂ���T�u�t�H���_�̕����̃}�N���t�@�C����Bas�t�@�C��
''             ���G�N�X�|�[�g����
''Arguments�Fnothing
''ReturnValue:nothing
''Note�F
''******************************************************************************
'
'Sub �T�u�t�H���_�܂ޕ����̃}�N���t�@�C����Bas�t�@�C�����G�N�X�|�[�g()
'
'  On Error GoTo Err_Trap
'
'  Call m_common.�}�N���J�n
'
'
'
'
'
'  MsgBox "Done!!�@�������I�����܂��B", vbInformation, "�����I��"
'
'  Call m_common.�}�N���I��
'
'  Exit Sub
'
'Err_Trap:
'
'  'When an error occurs, display the contents of the error in a message box.
'  If Err.Number <> 0 Then
'    '
'    Debug.Print Err.Number & " " & Err.Description
'    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error"
'
'    'Clear error
'    Err.Clear
'
'  End If
'
'
'
'End Sub
'
'
'
'
