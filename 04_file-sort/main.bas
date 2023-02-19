Attribute VB_Name = "main"
'###############################
''機能名：ファイル仕分けマクロ
''Author：okayasu jun
''作成日：2021/10/30
''更新日：2022/12/22
'COMMENT：
'###############################

'■実行条件
'処理種別
Public processType As String
'ファイル名中ブランクに置換する文字列
Public replaceTextToBlank As String
'#チェック処理
Sub exe()
    '実行
    Call main
    '完了通知
    MsgBox noticeCount & "件取得しました。"
End Sub
'#メイン処理
Function main()
    '共通初期化処理
    Set customSettingCurrentRange = initialize
    
    '独自初期化処理
    Call initializeInCustom(customSettingCurrentRange)
    
    'ファイルごとにチェック実行
    Call scanLoopWithFile(srcDirPath)
        
    '実行件数
    noticeCount = logWriteLine - 2
    
    '終了処理
    Call finally
End Function
'#機能独自初期化処理
Function initializeInCustom(customSettingCurrentRange As Range)
    
    processType = customSettingCurrentRange.Offset(0, 1).value
    replaceTextToBlank = customSettingCurrentRange.Offset(1, 1).value
    
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "元フォルダ"
    logSheet.Cells(1, 3) = "元ファイル名"
    logSheet.Cells(1, 4) = "先フォルダ"
    logSheet.Cells(1, 5) = "先ファイル名"
    logSheet.Cells(1, 6) = "処理種別"
    logSheet.Cells(1, 7) = "時刻"
End Function
'#対象の全ファイルを走査する。オプションに応じて再帰処理を行う。
Function scanLoopWithFile(argDirPath As String)
    'フォルダ内の最初のファイル名を取得
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '走査中ファイルが条件を通過するかどうか
        If isPassFile(argDirPath, currentFileName) Then
            'ファイル名を条件としたコピー or 移動
            Call fileMoveOrCopy(argDirPath, currentFileName)
        End If
        
        '次のファイル名を取り出す（なければブランク）
        currentFileName = Dir()
    Loop
    
    If recursiveFlag Then
        'フォルダ内のサブフォルダを順に取得
        For Each directory In objFSO.getfolder(argDirPath).SubFolders
            '再帰処理
            Call scanLoopWithFile(directory.path)
        Next
    End If
End Function
'移動元フォルダは再帰処理の場合、グローバル変数の値と異なる可能性があるため
'適宜引数から取得する
Function fileMoveOrCopy(argDirPath As String, currentFileName As String)

    '移動/コピー元、先、チェック用ファイル名格納用
    Dim srcFileName As String: srcFileName = currentFileName
    Dim distFileName As String

    '指定文字をブランクへ置換
    distFileName = Replace(currentFileName, replaceTextToBlank, "")
    
    If processType = "移動" Then
        '移動元　As 移動先
        Name argDirPath & "\" & srcFileName As distDirPath & "\" & distFileName
    ElseIf processType = "コピー" Then
        'コピー元, 移動先
        objFSO.CopyFile argDirPath & "\" & srcFileName, distDirPath & "\" & distFileName
    End If
    
    If logFlag Then
        'ログ記録
        logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
        logSheet.Cells(logWriteLine, 2) = argDirPath & "\"
        logSheet.Cells(logWriteLine, 3) = srcFileName
        logSheet.Cells(logWriteLine, 4) = distDirPath
        logSheet.Cells(logWriteLine, 5) = distFileName
        logSheet.Cells(logWriteLine, 6) = processType
        logSheet.Cells(logWriteLine, 7) = Format(Now(), "yyyy/mm/dd hh:mm:ss")
        logWriteLine = logWriteLine + 1
    End If

End Function
