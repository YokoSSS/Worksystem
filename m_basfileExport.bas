Attribute VB_Name = "m_basfileExport"
'Option Explicit
'
''******************************************************************************
''FunctionName:１つのファイルのBasファイルをエクスポートする
''Specifications：１つのファイルのBasファイルをエクスポートする
''Arguments：nothing
''ReturnValue:nothing
''Note：
''******************************************************************************
'Sub 一つのファイルのBasファイルをエクスポートする()
'
'Dim module                  As VBComponent      '// モジュール
'Dim moduleList              As VBComponents     '// VBAプロジェクトの全モジュール
'Dim extension                                   '// モジュールの拡張子
'Dim sPath                   As String           '// 処理対象ブックのパス
'Dim sFilePath               As String           '// エクスポートファイルパス
'Dim sFoldPath               As String           '// 処理対象ブックのフォルダパス
'Dim TargetBook              As Workbook         '// 処理対象ブックオブジェクト
'
'  On Error GoTo Err_Trap
'
'  Call m_common.マクロ開始
'
'  '処理対象ブックの選択
'  If GETファイル(sPath) = False Then Exit Sub
'
'  '処理対象ブックのフォルダ生成
'  sFoldPath = Mid(sPath, 1, Len(sPath) - 5)
'
'  If 指定したパスでフォルダ生成(sFoldPath) = False Then
'
'    Call m_common.マクロ終了
'
'    Exit Sub
'
'  End If
'
'  ThisWorkbook.Worksheets("main").Range("main_file1") = sFoldPath
'
'  '処理対象ブックのオブジェクト生成
'  Workbooks.Open stritems:=sPath
'
'  Set TargetBook = ActiveWorkbook
'
'  '処理対象ブックのモジュール一覧を取得
'  Set moduleList = TargetBook.VBProject.VBComponents
'
'  '// VBAプロジェクトに含まれる全てのモジュールをループ
'  For Each module In moduleList
'    '// クラス
'    If (module.Type = vbext_ct_ClassModule) Then
'        extension = "cls"
'    '// フォーム
'    ElseIf (module.Type = vbext_ct_MSForm) Then
'        '// .frxも一緒にエクスポートされる
'        extension = "frm"
'    '// 標準モジュール
'    ElseIf (module.Type = vbext_ct_StdModule) Then
'        extension = "bas"
'    '// その他
'    Else
'        '// エクスポート対象外のため次ループへ
'        GoTo CONTINUE
'    End If
'
'    '// エクスポート実施
'    sFilePath = sFoldPath & "\" & module.Name & "." & extension
'
'    Call module.Export(sFilePath)
'
'    '// 出力先確認用ログ出力
'    Debug.Print sFilePath
'
'CONTINUE:
'    Next module
'
'  'ブックを閉じる
'  TargetBook.Close SaveChanges:=False
'
'  'メモリ解放
'  Set TargetBook = Nothing
'
'  MsgBox "Done!!　処理を終了します。", vbInformation, "処理終了"
'
'  Call m_common.マクロ終了
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
'  Call m_common.マクロ終了
'
'End Sub
'
''******************************************************************************
''FunctionName:同フォルダ複数ファイルのBasファイルをエクスポートする
''Specifications：同じフォルダにある複数のマクロファイルのBasファイルを
''                エクスポートする
''Arguments：nothing
''ReturnValue:nothing
''Note：
''******************************************************************************
'Sub 同フォルダ複数ファイルのBasファイルをエクスポートする()
'
'Dim module                  As VBComponent      '// モジュール
'Dim moduleList              As VBComponents     '// VBAプロジェクトの全モジュール
'Dim extension                                   '// モジュールの拡張子
'Dim sPath                   As String           '// 処理対象ブックのパス
'Dim sFilePath               As String           '// エクスポートファイルパス
'Dim sFoldPath               As String           '// 処理対象ブックのフォルダパス
'Dim sFdPath                 As String           '// 処理対象ブックのフォルダパス
'Dim TargetBook              As Workbook         '// 処理対象ブックオブジェクト
'Dim FSO As New FileSystemObject
'Dim objFiles As File
'Dim objFolders As Folder
'Dim cntlist As Long
'
'  On Error GoTo Err_Trap
'
'  Call m_common.マクロ開始
'
'  cntlist = 1
'
'  '一括処理対象フォルダの選択
'  If GETフォルダ(sFdPath) = False Then Exit Sub
'
'  '指定フォルダにbasフォルダがなかったら作成し、あったら処理を強制終了する
'  If FolderExists(sFdPath & "\" & "bas") = True Then
'
'    MsgBox "指定したフォルダにbasフォルダがありました。" & vbCrLf & _
'        "処理を終了します。", vbCritical, "処理終了"
'
'    Call m_common.マクロ終了
'
'    Exit Sub
'
'  ElseIf FolderExists(sFdPath & "\" & "bas") = False Then
'
'    'basフォルダがなかったら作成
'    If 指定したパスでフォルダ生成(sFdPath & "\" & "bas") = False Then Exit Sub
'
'  End If
'
'  'ファイル名の取得
'  For Each objFiles In FSO.GetFolder(sFdPath).Files
'
'    Debug.Print "objFiles: " & objFiles
'
'    'マクロファイルを対象に処理をする
'    If FSO.GetExtensionName(objFiles.Name) = "xlsm" And _
'      objFiles.Name <> ThisWorkbook.Name And _
'       Not (objFiles.Name Like "*~$*") Then
'
'      'listシートにファイル情報を出力するための行カウント
'      cntlist = cntlist + 1
'
'      'リストに一括処理対象フォルダ内ファイルを書き出す
'      'If setFileList(searchPath, cntlist) = False Then Exit For
'
'      '処理対象ブックのフォルダ生成
'      sFoldPath = sFdPath & "\" & "bas" & "\" & FSO.GetBaseName(objFiles.Name)
'
'      Debug.Print "フォルダ: " & sFoldPath
'
'      If 指定したパスでフォルダ生成(sFoldPath) = False Then Exit Sub
'        'listシートにファイル情報を出力する
'
'        With ThisWorkbook.Worksheets("list")
'
'          'idn
'          .Cells(cntlist, Range("list_idn").Column) = cntlist - 1
'          'stritems
'          .Cells(cntlist, Range("list_stritems").Column) = objFiles.Name
'          'filepath
'          .Cells(cntlist, Range("list_filepath").Column) = list_failepath
'          '格納basフォルダpath
'          .Cells(cntlist, Range("list_baspath").Column) = list_failepath
'
'        End With
'
'        '処理対象ブックのオブジェクト生成
'        Workbooks.Open stritems:=objFiles
'
'        Set TargetBook = ActiveWorkbook
'
'        '処理対象ブックのモジュール一覧を取得
'        Set moduleList = TargetBook.VBProject.VBComponents
'
'        'VBAプロジェクトに含まれる全てのモジュールをループ
'        For Each module In moduleList
'          'クラス
'          If (module.Type = vbext_ct_ClassModule) Then
'              extension = "cls"
'          'フォーム
'          ElseIf (module.Type = vbext_ct_MSForm) Then
'              '.frxも一緒にエクスポートされる
'              extension = "frm"
'          '標準モジュール
'          ElseIf (module.Type = vbext_ct_StdModule) Then
'              extension = "bas"
'          'その他
'          Else
'              'エクスポート対象外のため次ループへ
'
'              GoTo CONTINUE
'
'          End If
'
'          'エクスポート実施
'          sFilePath = sFoldPath & "\" & module.Name & "." & extension
'
'        Call module.Export(sFilePath)
'
'        '出力先確認用ログ出力
'        Debug.Print sFilePath
'
'CONTINUE:
'        Next module
'
'        'ブックを閉じる
'        TargetBook.Close SaveChanges:=False
'
'        'ブックをbasへ移動させる
'        Name objFiles As sFdPath & "\" & "bas" & "\" & objFiles.Name
'
'        'メモリ解放
'        Set TargetBook = Nothing
'
'    End If
'
'
'  Next objFiles
'
'  MsgBox "Done!!　処理を終了します。", vbInformation, "処理終了"
'
'
'  Call m_common.マクロ終了
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
'  Call m_common.マクロ終了
'
'End Sub
'
''******************************************************************************
''FunctionName:同フォルダのサブフォルダ含む複数のマクロファイルのBasファイル
''             をエクスポートする
''Specifications：同じフォルダにあるサブフォルダの複数のマクロファイルのBasファイル
''             をエクスポートする
''Arguments：nothing
''ReturnValue:nothing
''Note：
''******************************************************************************
'
'Sub サブフォルダ含む複数のマクロファイルのBasファイルをエクスポート()
'
'  On Error GoTo Err_Trap
'
'  Call m_common.マクロ開始
'
'
'
'
'
'  MsgBox "Done!!　処理を終了します。", vbInformation, "処理終了"
'
'  Call m_common.マクロ終了
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
