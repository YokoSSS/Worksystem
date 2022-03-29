Attribute VB_Name = "Module1"
Option Explicit

Sub sample001()

On Error GoTo err_sample001

Dim str As String


    str = "<A>"


err_sample001:




End Sub

Sub sub_フォルダー内ファイルカウント()

    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With FSO.GetFolder("C:\")
        
        MsgBox .Files.Count & "個のファイルがあります", vbInformation
    
    End With

    Set FSO = Nothing

End Sub

Sub sub_ドライブサーチ()

    With CreateObject("Scripting.FileSystemObject")
        
        If .DriveExists("E") Then
            
            MsgBox "Eドライブが存在します", vbInformation
        
        Else
            
            MsgBox "Eドライブは存在しません", vbExclamation
        
        End If
    
    End With

End Sub

Sub sub_カレントフォルダの後ろに新しいフォルダ名を追加()

    ''カレントフォルダの後ろに新しいフォルダ名を追加します
    Dim FSO As Object, buf As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    buf = InputBox("新しいフォルダ名は？")
    
    MsgBox FSO.BuildPath(CurDir, buf)
    
    Set FSO = Nothing

End Sub



Sub sub_ファイルコピー()

    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ''C:\Tmp\フォルダのBook1.xlsxを、C:\Work\フォルダにコピーします
    FSO.CopyFile "C:\Tmp\Book1.xlsx", "C:\Work\"
    
    ''C:\Tmp\フォルダのBook1.xlsxを、C:\Work\フォルダにSample.xlsxという名前でコピーします
    FSO.CopyFile "C:\Tmp\Book1.xlsx", "C:\Work\Sample.xlsx"
    
    Set FSO = Nothing

End Sub

Sub sub_フォルダコピー()

    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ''C:\Tmp\フォルダのSubフォルダを、C:\Work\フォルダにコピーします
    FSO.CopyFolder "C:\Tmp\Sub", "C:\Work\"
    
    Set FSO = Nothing

End Sub

Sub sub_サブフォルダ作成()

    ''C:\Work\フォルダにSubフォルダを作成します。
    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FSO.CreateFolder "C:\Work\Sub"
    
    Set FSO = Nothing

End Sub


Sub sub_フォルダにユーザーが指定した名前のフォルダを作成()

    ''C:\Work\フォルダにユーザーが指定した名前のフォルダを作成します。
    Dim FSO As Object, buf As String, Result As String
    
    buf = InputBox("C:\Workに作成するフォルダ名を入力してください")
    
    If buf = "" Then Exit Sub
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    Result = FSO.CreateFolder("C:\Work\" & buf)
    
    Set FSO = Nothing
    
    If Err = 0 Then
        
        MsgBox Result & vbCrLf & "を作成しました", vbInformation
    
    Else
        
        MsgBox Err.Description, vbExclamation
    
    End If
  
End Sub
