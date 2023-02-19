Attribute VB_Name = "main"
'###############################
'機能名：Excel情報集約マクロ
'Author：okayasu jun
'作成日：2021/10/03
'更新日：2023/02/19
'COMMENT：2023/02からGitHubで管理します。
'###############################

'■実行条件
'ファイル名転記先アドレス
Public fileNameCopyAddress As String
'ファイル名をシート名にするか設定
Public sheetNameFromFileNameFlag As Boolean
'ファイル名中ブランクに置換する文字列
Public replaceTextToBlank As String
'生成ブックパス
Dim newBookPath As String
'生成ブック
Dim newBook As Workbook
'集約シート情報配列
Dim copySheetArray As Variant
'シートタイプ（名前/番号）
Dim copyType As String
'#実行
Sub exe()
    '実行
    Call main
    '完了通知
    MsgBox noticeCount & "シートを集約しました。" & vbCrLf & newBookPath & "に格納しています。"
End Sub
'#メイン処理
Function main()
    '共通初期化処理（Module2）
    Set customSettingCurrentRange = initialize
    
    '独自初期化処理（Module1）
    Call initializeInCustom(customSettingCurrentRange)
    
    'ファイルごとにチェック実行
    Call scanLoopWithFile(srcDirPath)
    
    '集約結果フォルダの保存&クローズ
    newBook.SaveAs newBookPath
    newBook.Close SaveChanges:=False
        
    '実行件数
    noticeCount = logWriteLine - 2
    
    '終了処理
    Call finally
End Function
'#機能独自初期化処理
Function initializeInCustom(customSettingCurrentRange As Range)

    'シート名、シートNoの条件取得
    Dim sheetNameArrayPerSheet As Variant: sheetNameArrayPerSheet = Split(customSettingCurrentRange.Offset(0, 1).value, ",")
    Dim sheetNoArrayPerSheet As Variant: sheetNoArrayPerSheet = Split(customSettingCurrentRange.Offset(0, 6).value, ",")
    
    '集約対象シート情報
    copySheetArray = IIf(UBound(sheetNameArrayPerSheet) >= 0, sheetNameArrayPerSheet, sheetNoArrayPerSheet)
    copyType = IIf(UBound(sheetNameArrayPerSheet) >= 0, "name", "no")

    'ファイル名を転記するアドレス
    fileNameCopyAddress = getBottomEndRange(customSettingCurrentRange, 1).Offset(0, 1).value
    'ファイル名をシート名にするか設定
    sheetNameFromFileNameFlag = getBottomEndRange(customSettingCurrentRange, 1).Offset(1, 1).value = "する"
    'ファイル名中ブランクに置換する文字列
    replaceTextToBlank = getBottomEndRange(customSettingCurrentRange, 1).Offset(2, 1).value


    '成果保存用新規ブック生成
    time = Format(Now(), "yyyy-mm-dd-hh-mm-ss")
    newBookPath = ThisWorkbook.path & "\" & "集約結果_" & time & ".xlsx"
    Set newBook = Workbooks.Add

    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "フォルダ"
    logSheet.Cells(1, 3) = "ファイル"
    logSheet.Cells(1, 4) = "元シート名"
    logSheet.Cells(1, 5) = "先シート名"
    logSheet.Cells(1, 6) = "コピー根拠"
    logSheet.Cells(1, 7) = "時刻"
End Function
'#対象の全ファイルを走査する。オプションに応じて再帰処理を行う。
Function scanLoopWithFile(argDirPath As String)
    'フォルダ内の最初のファイル名を取得
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    '順次開いていくbookの格納先
    Dim wb As Workbook
    'コピー元のシート名
    Dim srcSheetName As String
    'コピー先のシート名
    Dim distSheetName As String
    'コピー元ブックの拡張子
    Dim srcBookExt As String

    Do While currentFileName <> ""
    
        '走査中ファイルが条件を通過するかどうか
        If isPassFile(argDirPath, currentFileName) Then

            Set wb = Workbooks.Open(fileName:=argDirPath & "\" & currentFileName, UpdateLinks:=0)
            For Each target In copySheetArray

                srcBookExt = objFSO.GetExtensionName(argDirPath & "\" & currentFileName)
            
                'シート名・番号いずれかで処理実施。
                If copyType = "name" And isExistCheckToSheet(wb, target) Then
                    'シート名指定
                    wb.Sheets(target).Copy After:=newBook.Sheets(newBook.Sheets.count)
                    srcSheetName = wb.Sheets(target).Name
                ElseIf copyType = "no" And isExistCheckToSheet(wb, target) Then
                    'シート番号指定
                    wb.Sheets(CInt(target)).Copy After:=newBook.Sheets(newBook.Sheets.count)
                    srcSheetName = wb.Sheets(CInt(target)).Name
                End If

                'シート名をファイル名にする処理を通らない場合のため
                distSheetName = newBook.Sheets(newBook.Sheets.count).Name

                'ファイル名をコピー先シートのどこかのセルに転記する場合
                If isExistCheckToSheet(wb, target) And fileNameCopyAddress <> "" Then
                    newBook.Sheets(newBook.Sheets.count).Range(fileNameCopyAddress) = wb.Name
                End If

                'シート名をファイル名にする場合
                If isExistCheckToSheet(wb, target) And sheetNameFromFileNameFlag Then
                    distSheetName = Replace(wb.Name, replaceTextToBlank, "")
                    '拡張子削除（ドットも消す）
                    distSheetName = Replace(distSheetName, "." & srcBookExt, "")
                    
                    'シート名の禁止文字を削除
                    distSheetName = replaceTabooStrWithSheetName(distSheetName)
                    
                    '要素数が1（Ubound上は0）の場合、1ファイルにつき1シートになるのでファイル名のみとする
                    If UBound(copySheetArray) = 0 Then
                        newBook.Sheets(newBook.Sheets.count).Name = distSheetName
                    Else
                        'ファイル名 + シート名の形式。一意にするため
                        distSheetName = distSheetName & "-" & srcSheetName
                        newBook.Sheets(newBook.Sheets.count).Name = distSheetName
                    End If
                End If

                'ログ記録
                If isExistCheckToSheet(wb, target) And logFlag Then
                    logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
                    logSheet.Cells(logWriteLine, 2) = argDirPath & "\"
                    logSheet.Cells(logWriteLine, 3) = currentFileName
                    logSheet.Cells(logWriteLine, 4) = srcSheetName
                    logSheet.Cells(logWriteLine, 5) = distSheetName
                    logSheet.Cells(logWriteLine, 6) = target
                    logSheet.Cells(logWriteLine, 7) = Format(Now(), "yyyy/mm/dd hh:mm:ss")
                    logWriteLine = logWriteLine + 1
                Else
                    'コピー元ブックにシートがなく、コピーしていない場合もログに残すかどうか
                End If

            Next
            wb.Close SaveChanges:=False
            '急な処理落ち対策として保存
            Debug.Print newBookPath
            newBook.SaveAs newBookPath

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
    
    '列幅調整
    logSheet.Columns("A:G").AutoFit
    
End Function
