Attribute VB_Name = "Module1"
Option Explicit

Sub sample001()

On Error GoTo err_sample001

Dim str As String


    str = "<A>"


err_sample001:




End Sub

Sub sub_�t�H���_�[���t�@�C���J�E���g()

    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With FSO.GetFolder("C:\")
        
        MsgBox .Files.Count & "�̃t�@�C��������܂�", vbInformation
    
    End With

    Set FSO = Nothing

End Sub

Sub sub_�h���C�u�T�[�`()

    With CreateObject("Scripting.FileSystemObject")
        
        If .DriveExists("E") Then
            
            MsgBox "E�h���C�u�����݂��܂�", vbInformation
        
        Else
            
            MsgBox "E�h���C�u�͑��݂��܂���", vbExclamation
        
        End If
    
    End With

End Sub

Sub sub_�J�����g�t�H���_�̌��ɐV�����t�H���_����ǉ�()

    ''�J�����g�t�H���_�̌��ɐV�����t�H���_����ǉ����܂�
    Dim FSO As Object, buf As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    buf = InputBox("�V�����t�H���_���́H")
    
    MsgBox FSO.BuildPath(CurDir, buf)
    
    Set FSO = Nothing

End Sub



Sub sub_�t�@�C���R�s�[()

    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ''C:\Tmp\�t�H���_��Book1.xlsx���AC:\Work\�t�H���_�ɃR�s�[���܂�
    FSO.CopyFile "C:\Tmp\Book1.xlsx", "C:\Work\"
    
    ''C:\Tmp\�t�H���_��Book1.xlsx���AC:\Work\�t�H���_��Sample.xlsx�Ƃ������O�ŃR�s�[���܂�
    FSO.CopyFile "C:\Tmp\Book1.xlsx", "C:\Work\Sample.xlsx"
    
    Set FSO = Nothing

End Sub

Sub sub_�t�H���_�R�s�[()

    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ''C:\Tmp\�t�H���_��Sub�t�H���_���AC:\Work\�t�H���_�ɃR�s�[���܂�
    FSO.CopyFolder "C:\Tmp\Sub", "C:\Work\"
    
    Set FSO = Nothing

End Sub

Sub sub_�T�u�t�H���_�쐬()

    ''C:\Work\�t�H���_��Sub�t�H���_���쐬���܂��B
    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FSO.CreateFolder "C:\Work\Sub"
    
    Set FSO = Nothing

End Sub


Sub sub_�t�H���_�Ƀ��[�U�[���w�肵�����O�̃t�H���_���쐬()

    ''C:\Work\�t�H���_�Ƀ��[�U�[���w�肵�����O�̃t�H���_���쐬���܂��B
    Dim FSO As Object, buf As String, Result As String
    
    buf = InputBox("C:\Work�ɍ쐬����t�H���_������͂��Ă�������")
    
    If buf = "" Then Exit Sub
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    Result = FSO.CreateFolder("C:\Work\" & buf)
    
    Set FSO = Nothing
    
    If Err = 0 Then
        
        MsgBox Result & vbCrLf & "���쐬���܂���", vbInformation
    
    Else
        
        MsgBox Err.Description, vbExclamation
    
    End If
  
End Sub
