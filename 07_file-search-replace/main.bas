Attribute VB_Name = "main"
'###############################
'機能名：検索・置換v1.5
'Author：okayasu jun
'作成日：2022/12/05
'更新日：2022/12/25
'COMMENT：
'###############################

'■実行条件
'実行ケース１条件格納配列
Dim exe1Array As Variant
'実行ケース２条件格納配列
Dim exe2Array As Variant
'#処理
Sub exe()
    '実行
    Call main
    '完了通知
    MsgBox noticeCount & "件処理しました。"
End Sub
'#メイン処理
Function main()
    '共通初期化処理（Module2）
    Set customSettingCurrentRange = initialize
    
    '独自初期化処理（Module1）
    Call initializeInCustom(customSettingCurrentRange)
    
    Select Case Application.Caller
        Case "search1"
            Call scanLoopWithFile(srcDirPath)
        Case "replace1"
            Call scanLoopWithFile(srcDirPath)
        Case "get"
            Call scanLoopWithFile(srcDirPath)
        Case "search2"
            Call search2exe
        Case "replace2"
            Call replace2exe
    End Select
    
    '実行件数
    noticeCount = logWriteLine - 2
    
    '列幅調整
    logSheet.Columns("A:G").AutoFit
    
    '終了処理
    Call finally
End Function
'#機能独自初期化処理
Function initializeInCustom(customSettingCurrentRange As Range)
    '実行ケース１
    Dim exe1StartRange As Range
    Set exe1StartRange = customSettingCurrentRange.Offset(0, 1)
    exe1Array = IIf(IsEmpty(exe1StartRange.value), _
                    exe1StartRange, _
                    topSheet.Range(exe1StartRange, regionEndRange(exe1StartRange.Offset(0, 1))))

    '実行ケース２
    Dim exe2StartRange As Range
    Set exe2StartRange = getBottomEndRange(customSettingCurrentRange, 2).Offset(0, 1)
    exe2Array = IIf(IsEmpty(exe2StartRange.value), _
                        exe2StartRange, _
                        topSheet.Range(exe2StartRange, regionEndRange(exe2StartRange.Offset(0, 1), headerFlag:=True)))


    'ファイル取得用
    If Application.Caller = "get" Then
        'ファイル名書き出しに備えて
        Set customSettingCurrentRange = exe2StartRange
        '既存値のクリア（値のみ）
        topSheet.Range(customSettingCurrentRange, topSheet.Cells(getBottomEndRange(customSettingCurrentRange, 1).row, 5)).ClearContents
    End If

    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "フォルダ"
    logSheet.Cells(1, 3) = "ファイル名"
    logSheet.Cells(1, 4) = "検出・置換情報"
    logSheet.Cells(1, 5) = "文字コード"
    logSheet.Cells(1, 6) = "実行契機"
    logSheet.Cells(1, 7) = "時刻"
End Function
'#対象の全ファイルを走査する。オプションに応じて再帰処理を行う。
Function scanLoopWithFile(argDirPath As String)
    'フォルダ内の最初のファイル名を取得
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '走査中ファイルが条件を通過するかどうか
        If isPassFile(argDirPath, currentFileName) Then
            Select Case Application.Caller
                Case "search1"
                    '汎用検索処理
                    Call search1exe(argDirPath, currentFileName)
                Case "replace1"
                    '汎用置換処理
                    Call replace1exe(argDirPath, currentFileName)
                Case "get"
                    'ファイル名取得
                    Call writeFileName(argDirPath, currentFileName)
            End Select
        End If
        
        '次のファイル名を取り出す（なければブランク）
        currentFileName = Dir()
    Loop
    
    If recursiveFlag Then
        'フォルダ内のサブフォルダを順に取得
        For Each directory In objFSO.getfolder(argDirPath).SubFolders
            '再帰処理
            Call scanLoopWithFile(directory.Path)
        Next
    End If
    
End Function
'#実行ケース１の検索処理
Function search1exe(argDirPath As String, argFileName As String)
    '検索対象の配列はここで渡す（別口から渡すケースもあるため）
    Call mainSearch(argDirPath, argFileName, exe1Array)

End Function
'検索処理を実施する
Function mainSearch(argDirPath As String, argFileName As String, searchArray As Variant)
    'ファイルのフルパス
    Dim filePath As String: filePath = argDirPath & "\" & argFileName

    Dim findedAddress As String

    If InStr(objFSO.getFile(filePath).Type, "Excel") > 0 Then
        'Excel系ファイル
        'ブック中から検索する
        findedAddress = isInExcelFile(filePath, searchArray)
        If findedAddress <> "" Then
            'ログ出力
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = argFileName
            logSheet.Cells(logWriteLine, 4) = findedAddress
            logSheet.Cells(logWriteLine, 5) = "-"
            logSheet.Cells(logWriteLine, 6) = Application.Caller
            time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    Else
        'テキスト系ファイルを想定
        findedAddress = isInTextFile(filePath, searchArray)
        If findedAddress <> "" Then
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = argFileName
            logSheet.Cells(logWriteLine, 4) = findedAddress
            logSheet.Cells(logWriteLine, 5) = judgeFileCharSet(argDirPath & "\" & argFileName)
            logSheet.Cells(logWriteLine, 6) = Application.Caller
            time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    End If
End Function
'#実行ケース１の置換処理
Function replace1exe(argDirPath As String, argFileName As String)

    '置換対象の配列はここで渡す（別口から渡すケースもあるため）
    Call mainReplace(argDirPath, argFileName, exe1Array)

End Function
'置換処理を実施する
Function mainReplace(argDirPath As String, argFileName As String, replaceArray As Variant)
    'ファイルのフルパス
    Dim filePath As String: filePath = argDirPath & "\" & argFileName
    
    Dim findedAddress As String
    
    If InStr(objFSO.getFile(filePath).Type, "Excel") > 0 Then
        'Excel系ファイル
        'ブック中から検索する
        findedAddress = replaceInExcelFile(filePath, replaceArray)
        If findedAddress <> "" Then
            'ログ出力
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = argFileName
            logSheet.Cells(logWriteLine, 4) = findedAddress
            logSheet.Cells(logWriteLine, 5) = "-"
            logSheet.Cells(logWriteLine, 6) = Application.Caller
            time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    Else
        'テキスト系ファイルを想定
        findedAddress = replaceInTextFile(filePath, filePath, replaceArray)
        If findedAddress <> "" Then
            'ログ出力
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = argFileName
            logSheet.Cells(logWriteLine, 4) = findedAddress
            logSheet.Cells(logWriteLine, 5) = judgeFileCharSet(argDirPath & "\" & argFileName)
            logSheet.Cells(logWriteLine, 6) = Application.Caller
            time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    End If
End Function
'#渡されたファイル名を所定のセルに書き出す。セルはインクリメントする
Function writeFileName(argDirPath As String, currentFileName As String)
    customSettingCurrentRange = argDirPath
    customSettingCurrentRange.Offset(0, 1) = currentFileName
    Set customSettingCurrentRange = customSettingCurrentRange.Offset(1, 0)
    logWriteLine = logWriteLine + 1
End Function
'#実行ケース２の検索処理
Function search2exe()
    Dim dirPath As String
    Dim fileName As String
    
    Dim searchArray As Variant
    
    For i = LBound(exe2Array) To UBound(exe2Array)
        dirPath = exe2Array(i, 1)
        fileName = exe2Array(i, 2)
        
        '実行ケース２の検索配列を実行ケース１のように変換する
        searchArray = convertArrayFrom2dTo2d(exe2Array, CInt(i), 3)
        
        If dirPath <> "" And fileName <> "" Then
            Call mainSearch(dirPath, fileName, searchArray)
        End If
    Next
End Function
'#実行ケース２の置換処理
Function replace2exe()
    Dim dirPath As String
    Dim fileName As String

    Dim replaceArray As Variant
    
    For i = LBound(exe2Array) To UBound(exe2Array)
        dirPath = exe2Array(i, 1)
        fileName = exe2Array(i, 2)
        
        '実行ケース２の置換配列を実行ケース１のように変換する
        replaceArray = convertArrayFrom2dTo2d(exe2Array, CInt(i), 3)

        If dirPath <> "" And fileName <> "" Then
            Call mainReplace(dirPath, fileName, replaceArray)
        End If
        
    Next
End Function
'#説明：二次元配列を特定行の特定列で二次元配列に変換する
'#引数：srcArray:配列、argReadRow:二次元の特定行、argReadStartColumn:二次元目の開始インデックス
'#戻値：一次元配列
Function convertArrayFrom2dTo2d(srcArray As Variant, argReadRow As Integer, argReadStartColumn As Integer)
    Dim returnArray() As Variant
    '要素数の設定
    ReDim returnArray(UBound(srcArray, 2), 3)
    Dim count As Integer: count = 0
    
    '二次元目についてループを行う
    For i = argReadStartColumn To UBound(srcArray, 2) Step 2

        If srcArray(argReadRow, i) <> "" Then
            '検索値
            returnArray(count, 1) = srcArray(argReadRow, i)
            '置換値
            returnArray(count, 3) = srcArray(argReadRow, i + 1)
            count = count + 1
        End If
    Next
    convertArrayFrom2dTo2d = returnArray
End Function
