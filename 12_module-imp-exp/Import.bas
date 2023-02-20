Attribute VB_Name = "Import"
'#敬意/謝意：https://vbabeginner.net/bulk-import-of-standard-modules/
'参照設定：Microsoft Visual Basic for Application Extensibilly 5.3　を追加
'参照設定：Microsoft Scripting Runtime　を追加
'アクティブブック配下にある全モジュールをimportする
Sub ImportAll()
    On Error Resume Next
    '対象ブック（■ここはエクスプローラーだすか？）
    Dim TargetBook As Workbook
    Set TargetBook = ActiveWorkbook
    'インポート対象ファイルパス
    Dim importDirPath As String
    importDirPath = TargetBook.Path
    'ファイル操作オブジェクト
    Dim objFSO As New FileSystemObject
    'モジュール名配列
    Dim modulePathArray() As String
    'モジュール（ループで使用する都合でVariant型必須）
    'Dim modulePath  As Variant
    'モジュール拡張子
    Dim extension As String
    'ログ書き出し行
    Dim logWriteLine As Integer: logWriteLine = 2
    'ユーザ返答
    Dim response As VbMsgBoxResult
    
    response = MsgBox("同名のモジュールは上書きします。よろしいですか？", vbOKCancel, "上書き確認")
    If response <> vbOK Then
        Exit Sub
    End If
    
    'logシートの初期化
    Call logSetUp
    
    '配列要素数指定
    ReDim modulePathArray(0)
    
    '対象フォルダ配下の全モジュールのファイルパスを収集
    Call searchAllFile(importDirPath, modulePathArray)
    
    '全モジュールパスをループ
    For Each importFilePath In modulePathArray
        
        '拡張子を小文字で取得
        extension = LCase(objFSO.GetExtensionName(importFilePath))
        
        '拡張子がcls、frm、basのいずれかの場合
        If (extension = "cls" Or extension = "frm" Or extension = "bas") Then
            '同名モジュールを削除
            Call TargetBook.VBProject.VBComponents.Remove(TargetBook.VBProject.VBComponents(objFSO.GetBaseName(importFilePath)))
            'モジュールを追加
            Call TargetBook.VBProject.VBComponents.Import(importFilePath)
        
            'ログ出力
            ThisWorkbook.Worksheets("log").Cells(logWriteLine, 1) = logWriteLine - 1
            ThisWorkbook.Worksheets("log").Cells(logWriteLine, 2) = importFilePath
            ThisWorkbook.Worksheets("log").Cells(logWriteLine, 3) = "import"
            ThisWorkbook.Worksheets("log").Cells(logWriteLine, 4) = Now()
            logWriteLine = logWriteLine + 1
        End If
    Next
    
    '列幅調整
    ThisWorkbook.Worksheets("log").Columns("A:D").AutoFit
End Sub
'指定フォルダ配下の全ファイルパスを取得する
'argImportDirPath:走査対象ルートフォルダパス、argModulePathArray():ファイルパス配列
Function searchAllFile(argImportDirPath As String, argModulePathArray() As String)
    'ファイル操作オブジェクト
    Dim objFSO As New FileSystemObject
    '走査対象フォルダ
    Dim dir As Folder
    '走査対象サブフォルダ
    Dim subDir As Folder
    'ファイル
    Dim file As file
    '配列インデックス
    Dim i As Integer: i = 0
    
    If Not objFSO.FolderExists(argImportDirPath) Then
        'フォルダがなければ終了
        Exit Function
    End If
    
    '処理対象フォルダの取得
    Set dir = objFSO.GetFolder(argImportDirPath)
    
    'サブフォルダを再帰処理
    For Each subDir In dir.SubFolders
        Call searchAllFile(subDir.Path, argModulePathArray)
    Next
    
    'パス配列の要素数を取得
    i = UBound(argModulePathArray)
    
    '走査中フォルダ内のファイルを取得
    For Each file In dir.Files
    
        '要素がすでにあるかどうか。あればTrueを返す
        If (i <> 0 Or argModulePathArray(i) <> "") Then
            i = i + 1
            '要素値を保持したまま要素数を増加
            ReDim Preserve argModulePathArray(i)
        End If
        
        'ファイルパスを配列に格納（ここでは拡張子を限定しない）
        argModulePathArray(i) = file.Path
    Next
End Function
'ログシートの初期化
Function logSetUp()
    ThisWorkbook.Worksheets("log").Cells.Clear
    ThisWorkbook.Worksheets("log").Cells(1, 1) = "No"
    ThisWorkbook.Worksheets("log").Cells(1, 2) = "ファイル名"
    ThisWorkbook.Worksheets("log").Cells(1, 3) = "処理種別"
    ThisWorkbook.Worksheets("log").Cells(1, 4) = "実行時刻"
End Function
