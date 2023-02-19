Attribute VB_Name = "main"
'###############################
'機能名：ファイルリネームマクロ
'Author：okayasu jun
'作成日：2022/04/17
'更新日：2023/02/19
'COMMENT：
'###############################

'■実行ケース１
'連番付与
Public serialNoFlag As Boolean
'接頭辞
Public prefix As String
'接尾辞
Public suffix As String
'置換条件
Public replaceConditionArray As Variant
'■実行ケース２
'置換リスト
Public renameListArray As Variant
'リストの開始セル
Public renameFileStartRange As Range
'リネーム前ファイル名書き込み対象セル
Public renameWriteRange As Range
'番号付与用
Public serialNo As Integer
'#チェック処理
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
        Case "replace"
            '実行ケース１のファイル名置換処理
            Call scanLoopWithFile(srcDirPath, "replace")
            logSheet.Select
        Case "get"
            '実行ケース２のファイル名取得処理
            Call scanLoopWithFile(srcDirPath, "get")
        Case "serialNo"
            '実行ケース２の連番付与処理
            Call setSerialNo
        Case "rename"
            '実行ケース２のリネーム処理
            Call renameFile
            logSheet.Select
    End Select
        
    '実行件数
    noticeCount = logWriteLine - 2
    
    '終了処理
    Call finally
End Function
'#機能独自初期化処理
Function initializeInCustom(customSettingCurrentRange As Range)
    
    '■実行ケース１のパラメータ
    serialNoFlag = customSettingCurrentRange.Offset(0, 1) = "する"
    prefix = customSettingCurrentRange.Offset(1, 1)
    suffix = customSettingCurrentRange.Offset(2, 1)
    replaceConditionArray = Range(customSettingCurrentRange.Offset(3, 1), regionEndRange(customSettingCurrentRange.Offset(3, 1), False, 1))
    
    Set renameFileStartRange = getBottomEndRange(customSettingCurrentRange, 2)
    
    '■実行ケース２のパラメータ
    renameListArray = Range(renameFileStartRange.Offset(0, 1), regionEndRange(renameFileStartRange, False, 3))
    Set renameWriteRange = renameFileStartRange
    
    serialNo = 1
    
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "フォルダ"
    logSheet.Cells(1, 3) = "リネーム前"
    logSheet.Cells(1, 4) = "リネーム後"
    logSheet.Cells(1, 5) = "成功可否"
    logSheet.Cells(1, 6) = "ファイルの更新日時"
    logSheet.Cells(1, 7) = "処理時刻"
End Function
'#対象の全ファイルを走査する。オプションに応じて再帰処理を行う。
Function scanLoopWithFile(argDirPath As String, processType As String)
    'フォルダ内の最初のファイル名を取得
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '走査中ファイルが条件を通過するかどうか
        If isPassFile(argDirPath, currentFileName) Then
            If processType = "replace" Then
                'ファイル名を条件としたコピー or 移動
                Call replaceFileName(argDirPath, currentFileName)
                
            ElseIf processType = "get" Then
                '現在のファイル名書き出し処理
                Call writeFileName(argDirPath, currentFileName) '未実装
            End If
        End If
        
        '次のファイル名を取り出す（なければブランク）
        currentFileName = Dir()
    Loop
    
    If recursiveFlag Then
        'フォルダ内のサブフォルダを順に取得
        For Each directory In objFSO.getfolder(argDirPath).SubFolders
            '再帰処理
            Call scanLoopWithFile(directory.Path, processType)
        Next
    End If
End Function
'実行ケース１の使用
'移動元フォルダは再帰処理の場合、グローバル変数の値と異なる可能性があるため適宜引数から取得する
Function replaceFileName(argDirPath As String, currentFileName As String)

    'ブランク置換後文字
    Dim afterFileName As String
    '移動/コピー元、先、チェック用ファイル名格納用
    Dim srcFileName As String: srcFileName = currentFileName
    Dim distFileName As String ': distFileName = currentFileName
    Dim srcFilePath As String: srcFilePath = argDirPath & "\" & currentFileName
    Dim distFilePath As String
    Dim no As String
    
    'ファイルの拡張子（接尾辞付与のため）
    Dim currentFileExt As String: currentFileExt = objFSO.GetExtensionName(srcFilePath)
    
    '置換実施
    distFileName = replaceWithArray(replaceConditionArray, srcFileName)
    '接頭辞付与
    distFileName = prefix & distFileName
    '接尾辞付与
    distFileName = Replace(distFileName, "." & currentFileExt, suffix & "." & currentFileExt)
    'ファイル禁止文字削除
    distFileName = replaceTabooStrWithFileName(distFileName)
    
    '連番
    If serialNoFlag Then
        distFileName = Format(serialNo, "00") & "_" & distFileName
        serialNo = serialNo + 1
    End If
    
    'リネーム
    Name argDirPath & "\" & srcFileName As argDirPath & "\" & distFileName
    'ログ用
    distFilePath = argDirPath & "\" & distFileName
    
    If logFlag Then
        'ログ記録
        logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
        logSheet.Cells(logWriteLine, 2) = argDirPath & "\"
        logSheet.Cells(logWriteLine, 3) = srcFileName
        logSheet.Cells(logWriteLine, 4) = distFileName
        logSheet.Cells(logWriteLine, 5) = "=NOT(EXACT(C" & logWriteLine & ",D" & logWriteLine & "))"
        logSheet.Cells(logWriteLine, 6) = objFSO.getFile(distFilePath).DateLastModified
        logSheet.Cells(logWriteLine, 7) = Format(Now(), "yyyy/mm/dd hh:mm:ss")
        logWriteLine = logWriteLine + 1
    End If
End Function
'渡されたファイル名を所定のセルに書き出す。セル番地はインクリメントする（実行ケース２の使用想定）
'第一引数は使用していないが、今後の改修のため残している
Function writeFileName(argDirPath As String, currentFileName As String)
    renameWriteRange.Offset(0, 1) = argDirPath
    renameWriteRange.Offset(0, 2) = currentFileName
    Set renameWriteRange = renameWriteRange.Offset(1, 0)
    logWriteLine = logWriteLine + 1
End Function
'連番を付与する（実行ケース２の使用想定）
Function setSerialNo()
    'renameListArrayは件数を取得するためだけ（ちょっとださいな）。
    For i = LBound(renameListArray) To UBound(renameListArray)
        renameWriteRange.Offset(0, 6) = Format(serialNo, "00") & "_" & renameWriteRange.Offset(0, 6)
        'インクリメント
        Set renameWriteRange = renameWriteRange.Offset(1, 0)
        serialNo = serialNo + 1
        logWriteLine = logWriteLine + 1
    Next
End Function
'入力内容に従いリネームする（実行ケース２の使用想定）
Function renameFile()
    'リネーム前
    Dim srcDirName As String
    Dim srcFileName As String
    'リネーム後
    Dim distDirName As String
    Dim distFileName As String
    
    For i = LBound(renameListArray) To UBound(renameListArray)
        srcDirName = renameListArray(i, 1)
        srcFileName = renameListArray(i, 2)
        distDirName = renameListArray(i, 5)
        distFileName = renameListArray(i, 6)
        
        'ファイル禁止文字削除
        distFileName = replaceTabooStrWithFileName(distFileName)
        Name srcDirName & "\" & srcFileName As distDirName & "\" & distFileName

        If logFlag Then
            'ログ記録
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = srcDirName & "\"
            logSheet.Cells(logWriteLine, 3) = srcFileName
            logSheet.Cells(logWriteLine, 4) = distFileName
            logSheet.Cells(logWriteLine, 5) = "=NOT(EXACT(C" & logWriteLine & ",D" & logWriteLine & "))"
            logSheet.Cells(logWriteLine, 6) = objFSO.getFile(distDirName & "\" & distFileName).DateLastModified
            logSheet.Cells(logWriteLine, 7) = Format(Now(), "yyyy/mm/dd hh:mm:ss")
            logWriteLine = logWriteLine + 1
        End If
    Next
    
    '列幅調整
    logSheet.Columns("A:G").AutoFit
    
End Function
