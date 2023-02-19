Attribute VB_Name = "main"
'###############################
'機能名：テキストファイル置換コピーv2
'Author：okayasu jun
'作成日：2022/04/05
'更新日：2023/02/19
'COMMENT：
'###############################

'■実行条件
'シート指定
Dim replaceArray As Variant
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
    
    'ファイルごとにチェック実行
    Call scanLoopWithFile(srcDirPath)
        
    '実行件数
    noticeCount = logWriteLine - 2
    
    '終了処理
    Call finally
End Function
'#機能独自初期化処理
Function initializeInCustom(customSettingCurrentRange As Range)


    '■実行条件
    Dim replaceStartRange As Range
    Set replaceStartRange = customSettingCurrentRange.Offset(0, 1)
    Dim replaceEndRange As Range
    Set replaceEndRange = regionEndRange(replaceStartRange, headerFlag:=True, rightTimes:=1)
    replaceArray = IIf(replaceStartRange.value = "", replaceStartRange, topSheet.Range(replaceStartRange, replaceEndRange))
        
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "元フォルダ"
    logSheet.Cells(1, 3) = "元ファイル名"
    logSheet.Cells(1, 4) = "先フォルダ"
    logSheet.Cells(1, 5) = "先ファイル名"
    logSheet.Cells(1, 6) = "文字コード"
    logSheet.Cells(1, 7) = "時刻"
End Function
'#対象の全ファイルを走査する。オプションに応じて再帰処理を行う。
Function scanLoopWithFile(argDirPath As String)
    'フォルダ内の最初のファイル名を取得
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '走査中ファイルが条件を通過するかどうか
        If isPassFile(argDirPath, currentFileName) Then
        
            '置換ファイル生成
            Call replaceFileCopy(argDirPath, currentFileName)
        
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
    
    '列幅調整
    logSheet.Columns("A:G").AutoFit
    
End Function
'ファイルに置換処理を噛ませコピーする
Function replaceFileCopy(argDirPath As String, currentFileName As String)
    
    'コピー元ファイル
    Dim srcFileName As String: srcFileName = currentFileName
    
    'コピー先ファイル
    Dim distFileName As String
    
    '文字コードの判定結果
    Dim judgedCharSet As String
    
    'コピー元をフルパスで格納する
    Dim srcFilePath As String: srcFilePath = argDirPath & "\" & srcFileName
    
    'コピー先をフルパスで格納する
    Dim distFilePath As String
    
    '置換後文字列格納用
    Dim replacedContents As String
    
    '走査中ファイルが置換対象にあれば処理を行う
    For i = LBound(replaceArray) To UBound(replaceArray)
        
        If srcFileName = replaceArray(i, 1) Then
            'コピー先ファイル名
            distFileName = replaceTabooStrWithFileName(CStr(replaceArray(i, 2)))
            
            'コピー先ファイルパス
            distFilePath = distDirPath & "\" & distFileName
            
            'コピー元の文字コード検証
            judgedCharSet = judgeFileCharSet(srcFilePath)
            
            '置換前文字列取得
            replacedContents = getFileText(srcFilePath)
            
            '文字列置換（記載文すべて実施する）
            For j = 3 To UBound(replaceArray, 2) Step 2
                replacedContents = Replace(replacedContents, replaceArray(i, j), replaceArray(i, j + 1))
            Next
            
            If judgedCharSet = "UTF8" Then
                Call createTextFile(distFilePath, replacedContents, "UTF-8")
            ElseIf judgedCharSet = "SJIS" Then
                Call createTextFile(distFilePath, replacedContents, "SHIFT-JIS")
            Else
                judgedCharSet = "文字コード不正により未実施"
            End If
            
            'ログに記録
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = srcFileName
            logSheet.Cells(logWriteLine, 4) = distDirPath
            logSheet.Cells(logWriteLine, 5) = distFileName
            logSheet.Cells(logWriteLine, 6) = judgedCharSet
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    Next
End Function

