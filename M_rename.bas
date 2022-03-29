Attribute VB_Name = "M_rename"
Option Explicit
'******************************************************************************
Sub mainwork001()
Dim Fdnfullpath As String
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
        .PrintCommunication = False
        .EnableEvents = False
    End With
    
    Call CellDelete
    Call filenameget(Fdnfullpath)
        Debug.Print "( Fdnfullpath: " & Fdnfullpath & ")"
    Call extendget
    
    MsgBox "Processing Complete"
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .PrintCommunication = True
        .EnableEvents = True
    End With
End Sub
'******************************************************************************
Sub mainwork002()
Dim Fdnfullpath As String
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
        .PrintCommunication = False
        .EnableEvents = False
    End With
    
    Call FileRename(S_rename.Range("Fdnfullpath").Value)
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .PrintCommunication = True
        .EnableEvents = True
    End With

End Sub

'******************************************************************************
Sub CellDelete()
Dim lastrow As Long
Dim sr As Range
    
    lastrow = Cells(Rows.Count, S_rename.Range("NID").Column).End(xlUp).Row
    With S_rename
        If lastrow <> 2 Then
            Set sr = .Range("NID").Offset(1, 0)
            Debug.Print "(StartRange: " & sr.Address & ")"
            Debug.Print "(LastRange: " & _
                .Cells(lastrow, .Range("RenameComplete").Column).Address & ")"
'                Range(sr, .Cells(lastrow, .Range("extend").Column)).Clear
'                Range("Fdnfullpath").Clear
            .Range(sr, .Cells(lastrow, .Range("extend").Column)).ClearContents
'            .Range("Fdnfullpath").ClearContents
        Else
            Exit Sub
        End If
    
    End With
    
    Set sr = Nothing

End Sub

'******************************************************************************
Sub filenameget(Fdnfullpath As String)
Dim FileCollection As Object
Dim FileList As Variant
Dim cnt As Long
    
'    With Application.FileDialog(msoFileDialogFolderPicker)
'        If .Show = True Then
'            Fdnfullpath = .SelectedItems(1)
'            Debug.Print "(getfoldername: " & Fdnfullpath & ")"
'            S_rename.Range("Fdnfullpath").Value = Fdnfullpath & "\"
'        Else
'            MsgBox "Finish processing"
'            Exit Sub
'        End If
'    End With
    
    Fdnfullpath = S_rename.Range("Fdnfullpath").Value & "\"
    
    Set FileCollection = CreateObject("Scripting.FileSystemObject") _
            .GetFolder(Fdnfullpath).Files
    
    cnt = 0
    For Each FileList In FileCollection
        cnt = cnt + 1
        With S_rename
            .Range("NID").Offset(cnt, 0) = cnt
            .Range("CurrentFilename").Offset(cnt, 0) = FileList.Name
        End With
    Next
    
    Set FileCollection = Nothing


End Sub
'******************************************************************************
Sub FileRename(ByRef Fdnfullpath As String)
Dim lastrow As Long
Dim tarrg As Variant
Dim allr As Range

    Debug.Print "(getfolderpath: " & Fdnfullpath & ")"
    
    lastrow = Cells(Rows.Count, S_rename.Range("NID").Column).End(xlUp).Row - 2
    Debug.Print "(S_rename lastrow: " & lastrow & ")"

    Set allr = Range(S_rename.Range("CurrentFilename").Offset(1, 0), _
        S_rename.Range("CurrentFilename").Offset(lastrow, 0))
        Debug.Print "(allr: " & allr.Address & ")"

    For Each tarrg In allr
        
        Debug.Print "(tarrg: " & tarrg & ")"
        Debug.Print "(tarrg: " & tarrg.Offset(0, 1) & ")"
        If Dir(Fdnfullpath & tarrg) <> "" Then
            If tarrg <> "" And tarrg.Offset(0, 1) <> "" Then
                Name Fdnfullpath & tarrg As Fdnfullpath & tarrg.Offset(0, 1)
                tarrg.Offset(0, 2).Value = "Complete"
            
            ElseIf tarrg = "" Then
                tarrg.Offset(0, 2).Value = "Not Complete(Please enter before change file name.)"
            
            ElseIf tarrg.Offset(0, 1) = "" Then
                tarrg.Offset(0, 2).Value = "Not Complete(Please enter after change file name.)"
            
            End If
        Else
            tarrg.Offset(0, 2).Value = "Not Complete"
        End If
        
    Next
    
    MsgBox "Processing Complete"
    Set allr = Nothing
    
End Sub
'******************************************************************************
Sub extendget()
Dim i As Long
Dim tarrg As String
    With S_rename
        For i = 3 To .Cells(Rows.Count, 2).End(xlUp).Row
            tarrg = .Cells(i, 2).Value
            .Cells(i, 5).Value = Mid(tarrg, InStrRev(tarrg, "."), _
                Len(tarrg) - InStrRev(tarrg, ".") + 1)
        Next i
    End With
End Sub
'******************************************************************************
