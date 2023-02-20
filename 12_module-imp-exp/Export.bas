Attribute VB_Name = "Export"
'#敬意/謝意：https://vbabeginner.net/bulk-export-of-standard-modules/
'参照設定：Microsoft Visual Basic for Application Extensibilly 5.3　を追加
'アクティブブックのモジュールをブックと同フォルダにexportする
Sub ExportAll()
    'モジュール
    Dim module As VBComponent
    '全モジュール
    Dim moduleList As VBComponents
    'モジュールの拡張子
    Dim extension
    '処理対象ブックパス
    Dim bookPath As String
    'エクスポート対象ファイルパス
    Dim exportFilePath  As String
    '処理対象ブック
    Dim TargetBook As Workbook
    'ログ書き出し行
    Dim logWriteLine As Integer: logWriteLine = 2
    'ユーザ返答
    Dim response As VbMsgBoxResult
    'commonモジュールエクスポート対象フラグ
    Dim commonFlag As Boolean
    
    response = MsgBox("共通モジュールもエクスポートしますか？", vbYesNoCancel + vbQuestion)
    
    If response = vbCancel Then
        Exit Sub
    
    ElseIf response = vbNo Then
        commonFlag = False
        
    ElseIf response = vbYes Then
        commonFlag = True
        
    End If
    
    'logシートの初期化
    Call logSetUp
    
    If (Workbooks.Count = 1) Then
        '開いているブックが当ブックのみであればこれを対象とする
        Set TargetBook = ThisWorkbook
    Else
        '複数ブックを開いていれば処理実行時のアクティブブックを対象とする
        Set TargetBook = ActiveWorkbook
    End If
    
    '処理対象ブックのパスを取得
    bookPath = TargetBook.Path
    
    '処理対象ブックのモジュール一覧を取得
    Set moduleList = TargetBook.VBProject.VBComponents
    
    'VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
        
        If (module.Type = vbext_ct_ClassModule) Then
            'クラス
            extension = "cls"
        
        ElseIf (module.Type = vbext_ct_MSForm) Then
            'フォーム　※「.frx」も一緒にエクスポートされる
            extension = "frm"
        
        ElseIf (module.Type = vbext_ct_StdModule) Then
            '標準モジュール
            extension = "bas"
        Else
            'その他 エクスポート対象外のため次ループへ
            GoTo CONTINUE
        End If
        
        If module.Name = "common" And Not commonFlag Then
            '共通モジュールをエクスポートしない
            GoTo CONTINUE
        End If
        
        'エクスポート実施
        exportFilePath = bookPath & "\" & module.Name & "." & extension
        Call module.Export(exportFilePath)
        
        '出力先確認用ログ出力
        Debug.Print exportFilePath
        ThisWorkbook.Worksheets("log").Cells(logWriteLine, 1) = logWriteLine - 1
        ThisWorkbook.Worksheets("log").Cells(logWriteLine, 2) = exportFilePath
        ThisWorkbook.Worksheets("log").Cells(logWriteLine, 3) = "export"
        ThisWorkbook.Worksheets("log").Cells(logWriteLine, 4) = Now()
        logWriteLine = logWriteLine + 1
CONTINUE:
    Next
    
    '列幅調整
    ThisWorkbook.Worksheets("log").Columns("A:D").AutoFit
End Sub
'ログシートの初期化
Function logSetUp()
    ThisWorkbook.Worksheets("log").Cells.Clear
    ThisWorkbook.Worksheets("log").Cells(1, 1) = "No"
    ThisWorkbook.Worksheets("log").Cells(1, 2) = "ファイル名"
    ThisWorkbook.Worksheets("log").Cells(1, 3) = "処理種別"
    ThisWorkbook.Worksheets("log").Cells(1, 4) = "実行時刻"
End Function
