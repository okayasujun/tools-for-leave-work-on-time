Attribute VB_Name = "common"
'■基本情報
'実行シート
Public topSheet As Worksheet
'実行ログシート
Public logSheet As Worksheet
'ファイル操作オブジェクト
Public objFSO As Object
'元フォルダパス
Public srcDirPath As String
'先フォルダパス
Public distDirPath As String
'■共通オプション
'再帰処理フラグ
Public recursiveFlag As Boolean
'ログ記録フラグ
Public logFlag As Boolean
'対象ファイル拡張子
Public targetFileExtArray As Variant
'更新日時（FROM）
Public lastUpdateDateFrom As Date
'更新日時（TO）
Public lastUpdateDateTo As Date
'フォルダ構成再現フラグ
Public dirLevelCopyFlag As Boolean
'対象ファイル条件に拡張子も対象とするフラグ
Public extensionUseFlag As Boolean
'対象ファイル条件
Public targetFilterArray As Variant
'■その他
'ログ記録行
Public logWriteLine As Integer
'実行件数
Public noticeCount As Integer
'時間記録用
Public time As String
'機能独自の設定項目開始位置
Public customSettingCurrentRange As Range
'■機能独自の変数
'チェック条件格納配列
Public exe1Array As Variant
'ヘッダ行
Public headerRowNo As Integer
'#説明：初期処理
'#引数：なし
'#戻値：Range:機能セクションの最初の項目ラベルセル
Function initialize()

    '■実行環境情報
    '実行シート
    Set topSheet = ThisWorkbook.Sheets(1)
    'ログシート
    If isExistCheckToSheet(ActiveWorkbook, "log") Then
        Set logSheet = ThisWorkbook.Sheets(2)
    Else
        Set logSheet = Sheets.Add(After:=Sheets(1))
        logSheet.Name = "log"
    End If
    
    'ファイル操作オブジェクト
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    With topSheet
        '■基本情報
        Dim baseSectionInputStartRange As Range
        Set baseSectionInputStartRange = getBottomEndRange(.Cells(1, 3), 1)
        
        'コピー元フォルダ
        srcDirPath = baseSectionInputStartRange.Offset(0, 1).value
        
        'コピー先フォルダ
        distDirPath = baseSectionInputStartRange.Offset(1, 1).value

        '■共通オプション
        Dim commonOptionSectionInputStartRange As Range
        Set commonOptionSectionInputStartRange = getBottomEndRange(baseSectionInputStartRange, 2)
        
        '再帰処理フラグ
        recursiveFlag = commonOptionSectionInputStartRange.Offset(0, 1).value = "する"
        
        'ログ出力設定
        logFlag = commonOptionSectionInputStartRange.Offset(1, 1).value = "する"
        
        '対象ファイル形式
        targetFileExtArray = Split(commonOptionSectionInputStartRange.Offset(2, 1).value, ",")
        
        '更新日時（FROM）
        lastUpdateDateFrom = commonOptionSectionInputStartRange.Offset(3, 1).value
        
        '更新日時（TO）
        lastUpdateDateTo = commonOptionSectionInputStartRange.Offset(4, 1).value
        
        'フォルダ構成再現フラグ
        dirLevelCopyFlag = commonOptionSectionInputStartRange.Offset(5, 1).value = "する"
        
        '拡張子利用フラグ
        extensionUseFlag = commonOptionSectionInputStartRange.Offset(6, 1).value = "する"
        
        '条件（1:値、2,4,5,:不使用、3:条件種別、6:AND/OR）
        Dim targetFilterStartRange As Range
        Set targetFilterStartRange = commonOptionSectionInputStartRange.Offset(7, 1)
        Dim targetFilterEndRange As Range
        Set targetFilterEndRange = regionEndRange(targetFilterStartRange, rightTimes:=2)
        targetFilterArray = IIf(targetFilterStartRange.value = "", targetFilterStartRange, .Range(targetFilterStartRange, targetFilterEndRange))

        '■機能セクションの最初のラベルセルを返す
        Set initialize = getBottomEndRange(commonOptionSectionInputStartRange, 2)

    End With
    
    'ログの書き出し行
    logWriteLine = 2
    
    '記録時間
    time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
    
    '■実行動作条件
    '実行時の画面描画を静かにする
    Application.DisplayAlerts = False
    'トリガー式の自動マクロを起動させない
    Application.EnableEvents = False
    '画面停止
    Application.ScreenUpdating = False
End Function
'#説明：終了処理
'#引数：なし
'#戻値：なし
Function finally()
    '実行時の画面描画を元に戻す
    Application.DisplayAlerts = True
    'トリガー式自動マクロ封印を解除
    Application.EnableEvents = True
    '画面停止解除
    Application.ScreenUpdating = True
End Function
'#説明：指定されたシートのセルから指定回数分下移動したセルを返す
'#引数：argRange:開始位置セル、argTimes:移動回数
'#戻値：移動先のセル情報
Function getBottomEndRange(argRange As Range, argTimes As Integer)
    Dim returnRange As Range
    Set returnRange = argRange
    For i = 1 To argTimes
        Set returnRange = topSheet.Range(returnRange.Address).End(xlDown)
    Next
    '戻り値
    Set getBottomEndRange = returnRange
End Function
'#説明：指定されたシートのセルから指定回数分右移動したセルを返す
'#引数：argRange:開始位置セル、argTimes:移動回数
'#戻値：移動先のセル情報
Function getRightEndRange(argRange As Range, argTimes As Integer)
    Dim returnRange As Range
    Set returnRange = argRange
    For i = 1 To argTimes
        Set returnRange = topSheet.Range(returnRange.Address).End(xlToRight)
    Next
    '戻り値
    Set getRightEndRange = returnRange
End Function
'#説明：指定されたフォルダがなければ作成する
'#引数：argDirPath:フォルダパス（絶対パス）
'#戻値：なし
Function createDirectory(argDirPath As String)
    If Not objFSO.FolderExists(argDirPath) Then
        objFSO.CreateFolder (argDirPath)
    End If
End Function
'#説明：走査中ファイルが条件を通過するかどうかを返す（条件情報はグローバル変数からとる）
'#引数：argDirPath:フォルダパス（絶対パス）、argFileName:ファイル名
'#戻値：真偽値（true:通過、false:不適合）
Function isPassFile(argDirPath As String, argFileName As String)
    'ファイルフルパス
    Dim filePath As String: filePath = argDirPath & "\" & argFileName
    '拡張子の取得
    Dim fileExt As String: fileExt = objFSO.GetExtensionName(filePath)
    '拡張子を省いたファイル名
    Dim checkFileName As String: checkFileName = argFileName
    If Not extensionUseFlag Then
        checkFileName = Replace(argFileName, "." & fileExt, "")
    Else
    End If
    '対象ファイル条件に該当するか
    If isExistArray(targetFileExtArray, fileExt) _
        And isPassConditionCheck(checkFileName) _
        And isPassUpdateDate(filePath) Then
        '戻り値
        isPassFile = True
        Exit Function
    End If
    '戻り値
    isPassFile = False
End Function
'#説明：指定された値が指定された配列内に存在するかどうかを返す
'#引数：targetArray:走査配列、checkValue:検証値
'#戻値：真偽値（true:ある、false:ない）
Function isExistArray(targetArray As Variant, checkValue As String)
    isExistArray = False
    'UBoundの戻り値：-1は要素数0を示す
    If UBound(targetArray) = -1 Then
        isExistArray = True
        Exit Function
    End If
    
    For i = LBound(targetArray) To UBound(targetArray)
        If targetArray(i) = checkValue Then
            isExistArray = True
            Exit For
        End If
    Next
End Function
'#説明：指定されたファイル名を検査し、その結果を返す
'#引数：targetArray:走査配列、checkValue:検証値
'#戻値：真偽値（true:OK、false:NO）
Function isPassConditionCheck(argFileName As String)
    'チェック対象の値
    Dim checkValue As String
    'ループ中現在周回の検証結果真偽値
    Dim currentResult As Boolean
    '累計の検証結果真偽値
    Dim totalResult As Boolean
    '検証条件の種類
    Dim conditionType As String
    '条件の種類。AndかOrか
    Dim andor As String

    If IsEmpty(targetFilterArray) Then
        isPassConditionCheck = True
        Exit Function
    End If
    
    '最初の要素が空の場合、条件指定はなしと判断し、すべてtrueを返す
    If targetFilterArray(LBound(targetFilterArray, 1), 1) = "" Then
        isPassConditionCheck = True
        Exit Function
    End If
    
    'ループ内でも使うため一度変数にいれる
    Dim minIndex As Integer: minIndex = LBound(targetFilterArray)
    For i = minIndex To UBound(targetFilterArray)
    
        checkValue = targetFilterArray(i, 1)
        conditionType = targetFilterArray(i, 3)
        
        If checkValue = "" Or conditionType = "" Then
            GoTo continue
        End If
        
        If conditionType = "左記にファイル名が一致する" Then
            currentResult = (argFileName = checkValue)
            
        ElseIf conditionType = "左記をファイル名に含む" Then
            currentResult = InStr(argFileName, checkValue) > 0
            
        ElseIf conditionType = "左記をファイル名に含まない" Then
            currentResult = InStr(argFileName, checkValue) = 0
            
        ElseIf conditionType = "左記からファイル名が始まる" Then
            currentResult = isStartText(argFileName, checkValue)
            
        ElseIf conditionType = "左記でファイル名が終わる" Then
            currentResult = isEndText(argFileName, checkValue)
            
        ElseIf conditionType = "左記の正規表現に一致する" Then
            currentResult = isRegexpHit(argFileName, checkValue)
            
        End If
        
        If i = minIndex Then
            '最初の条件検証結果はそのまま初期値として設定する
            totalResult = currentResult
        Else
            '2つ目以降の条件はAND/ORと前回の結果によって導出する。引数の-1は前回の条件種別を取得したいから
            andor = CStr(targetFilterArray(i - 1, 6))
            totalResult = deriveTotalResult(totalResult, currentResult, andor)
        End If
continue:
    Next
    isPassConditionCheck = totalResult
End Function
'#説明：既存の真偽値と新たな真偽値から条件によって真偽値を導出する
'#引数：totalResult:既存の真偽値、currentResult:新しい真偽値、andor:条件種別（かつ・または）
'#戻値：真偽値（true:OK、false:NO）
Function deriveTotalResult(totalResult As Boolean, currentResult As Boolean, andor As String)
    If andor = "かつ" Or LCase(andor) = "and" Then
        totalResult = totalResult And currentResult
    ElseIf andor = "または" Or LCase(andor) = "or" Then
        totalResult = totalResult Or currentResult
    Else
        '未指定時はOR条件として扱う
        totalResult = totalResult Or currentResult
    End If
    '戻り値
    deriveTotalResult = totalResult
End Function
'#説明：文字列が指定の文字で始まるかどうかを返す
'#引数：largeText:検証対象の文字列、searchText:操作文字列
'#戻値：真偽値（true:始まる、false:始まらない）
Function isStartText(largeText As String, searchText As String)
    isStartsText = False
    If Len(searchText) > Len(largeText) Then
        '検証テキストが被検証テキストの長さを超える場合チェックが成立しない
        Exit Function
    End If
  
    If Left(largeText, Len(searchText)) = searchText Then
        isStartText = True
    End If
End Function
'#説明：文字列が指定の文字で終わるかどうかを返す
'#引数：largeText:検証対象の文字列、searchText:操作文字列
'#戻値：真偽値（true:終わる、false:終わらない）
Function isEndText(largeText As String, searchText As String)
    isEndText = False
    If Len(searchText) > Len(largeText) Then
        '検証テキストが被検証テキストの長さを超える場合チェックが成立しない
        Exit Function
    End If

    If Right(largeText, Len(searchText)) = searchText Then
        isEndText = True
    End If
End Function
'#説明：ファイルの最終更新日が条件かどうかを返す
'#引数：argFilePath:ファイルのフルパス
'#戻値：真偽値（true:条件内、false:条件外）
Function isPassUpdateDate(argFilePath As String)
    isPassUpdateDate = True
    Dim fileUpdateDate As Date: fileUpdateDate = objFSO.getFile(argFilePath).DateLastModified
    '未指定の場合「0」になるためこの条件
    If lastUpdateDateFrom <> 0 And Not lastUpdateDateFrom <= fileUpdateDate Then
        isPassUpdateDate = False
    End If
    If lastUpdateDateTo <> 0 And Not fileUpdateDate <= lastUpdateDateTo Then
        isPassUpdateDate = False
    End If
    
End Function
'#説明：ファイルの禁止文字を削除する
'#引数：fileName:ファイル名
'#戻値：禁止文字削除後の文字列
Function replaceTabooStrWithFileName(fileName As String)
    Dim tabooStringArray As Variant: tabooStringArray = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each taboo In tabooStringArray
        fileName = Replace(fileName, taboo, "")
    Next
    replaceTabooStrWithFileName = fileName
End Function
'#説明：Excelシートの禁止文字を削除する
'#引数：sheetName:シート名
'#戻値：禁止文字削除後の文字列
Function replaceTabooStrWithSheetName(sheetName As String)
    Dim tabooStringArray As Variant: tabooStringArray = Array(":", "：", "\", "￥", "?", "？", "[", "［", "]", "］", "/", "／", "*", "＊")
    For Each taboo In tabooStringArray
        sheetName = Replace(sheetName, taboo, "")
    Next
    replaceTabooStrWithSheetName = sheetName
End Function
'#説明：指定された値が指定された配列内の何番目に存在するかを返す
'#引数：targetArray:走査配列、checkValue:検査値
'#戻値：値が最初に出現インデックス
Function isExistArrayReturnIndex(targetArray As Variant, checkValue As String)
    isExistArrayReturnIndex = -1
    
    'UBoundの戻り値：-1は要素数0を示す。この場合、-1を返す
    If UBound(targetArray) = -1 Then
        isExistArrayReturnIndex = -1
        Exit Function
    End If
    
    For i = LBound(targetArray) To UBound(targetArray)
        If targetArray(i) = checkValue Then
            'ヒットしたインデックスを戻り値に設定
            isExistArrayReturnIndex = i
            Exit For
        End If
    Next
End Function
'#説明：渡されたファイルの文字コードを返す
'#引数：filePath:ファイルのフルパス
'#戻値：文字コード（SJIS、UTF8）
Function judgeFileCharSet(filePath As String)
    '判定のためにバイナリモードで取得する
    Dim bytCode() As Byte
    With CreateObject("ADODB.Stream")
        .Type = 1 'バイナリで開くため
        .Open
        .LoadFromFile filePath
        bytCode = .read
        .Close
    End With
    judgeFileCharSet = judgeCode(bytCode)
End Function
'#説明：二次元配列を特定の列（2次元目の配列）で一次元配列に変換する
'#引数：srcArray:二次元配列、argCol:二次元目のインデックス
'#戻値：一次元配列
Function convertArrayFrom2dTo1d(srcArray As Variant, argCol As Integer)
    Dim returnArray As Variant
    For i = LBound(srcArray) To UBound(srcArray)
        If i = LBound(srcArray) Then
            '最初の要素は「Preserve」を使わないからこの分岐
            ReDim returnArray(0)
            returnArray(0) = srcArray(i, argCol)
        Else
            '配列の要素数を1増やす
            ReDim Preserve returnArray(UBound(returnArray) + 1)
            '増やした要素に値を格納する
            returnArray(UBound(returnArray)) = srcArray(i, argCol)
        End If
    Next
    convertArrayFrom2dTo1d = returnArray
End Function
'#説明：全シートを走査する。詳細処理は機能側で実装する
'#引数：wb:ブック、exeSheet:処理対象のシート（シート名/番号可。カンマ区切りで複数指定も可）
'#戻値：なし
Function scanWithAllSheets(filePath As String, exeSheet As String)
    Dim wb As Workbook
    'ここでブックを開く
    Set wb = Workbooks.Open(fileName:=filePath, UpdateLinks:=0)

    '走査用
    Dim ws As Worksheet
    Dim exeSheetNo As Integer
    Dim exeSheetName As String
    Dim exeSheetArray As Variant
    exeSheetArray = Split(exeSheet, ",")
    'TODO:ここ「exeSheet」の値に「全」が入ってたら取捨選択するループを使わず、愚直にブックのすべてを対象にするロジック組もう
    
    For i = LBound(exeSheetArray) To UBound(exeSheetArray)
        If IsNumeric(exeSheetArray(i)) Then
            'シートを番号で指定
            exeSheetNo = CInt(exeSheetArray(i))
            If isExistCheckToSheet(wb, exeSheetNo) Then
                Set ws = wb.Worksheets(exeSheetNo)
                '独自処理にシートを渡す
                Call customProcess(ws)
            End If
        Else
            'シートを名前で指定
            exeSheetName = CStr(exeSheetArray(i))
            If isExistCheckToSheet(wb, exeSheetName) Then
                Set ws = wb.Worksheets(exeSheetName)
                '独自処理にシートを渡す
                Call customProcess(ws)
            End If
            
        End If
    Next
    
    wb.Close SaveChanges:=False
End Function
'#説明：指定されたシートの使用範囲最終セルを取得する
'#引数：ws:シート
'#戻値：使用範囲最終セルのアドレス
Function usedLastRange(ws As Worksheet)
    Dim addressArray As Variant
    addressArray = Split(ws.UsedRange.Address, ":")
    If UBound(addressArray) = 0 Then
        '単一セルの場合
        Set lastRange = ws.Range(Split(ws.UsedRange.Address, ":")(0))
    Else
        '複数セルに渡す場合
        Set lastRange = ws.Range(Split(ws.UsedRange.Address, ":")(1))
    End If
    Set usedLastRange = lastRange
End Function
'#説明：指定されたセルを起点とした範囲の右下セルを取得する。右移動にヘッダを使うか、右移動を何回行うかを指定できる
'#引数：startRange:開始位置セル、headerFlag:ヘッダ行で右移動をするか、rightTimes:右移動回数
'#戻値：始点から取得できた右下のセル
Function regionEndRange(startRange As Range, Optional headerFlag As Boolean = False, Optional rightTimes As Integer = 1)
        Dim rightEndRange As Range
        Dim tempRange As Range
        Set tempRange = startRange
        'ヘッダフラグがTRUEの場合、ヘッダで右方向の最終列を取得する
        If headerFlag Then
            Set tempRange = startRange.Offset(-1, 0)
        End If
        
        Set rightEndRange = getRightEndRange(tempRange, rightTimes)
        Dim bottomEndRange As Range
        If getBottomEndRange(startRange, 1).value <> "" Then
            Set bottomEndRange = getBottomEndRange(startRange, 1)
        Else
            Set bottomEndRange = startRange
        End If
        '引数の開始セルから行方向、列方向にずらしたセルを戻り値に設定
        Set regionEndRange = startRange.Offset(bottomEndRange.row - startRange.row, rightEndRange.Column - startRange.Column)
End Function
'#説明：指定されたファイルの中身を取得する
'#引数：fullPath:ファイルのフルパス
'#戻値：テキストファイルの内容
Function getFileText(fullPath As String)
    '文字コード判定
    Dim charset As String: charset = judgeFileCharSet(fullPath)
    If charset = "UTF8" Then
        charset = "UTF-8"
    ElseIf charset = "SJIS" Then
        charset = "SHIFT-JIS"
    End If
    
    With CreateObject("ADODB.Stream")
        .charset = charset
        .Open
        .LoadFromFile fullPath
        '戻り値
        getFileText = .ReadText
        .Close
    End With
End Function
'#説明：指定されたパスからフォルダもしくはファイル名を取得する。
'#引数：argFilePath:ファイルのフルパス、witch:1ならフォルダ、それ以外ならファイル名
'#戻値：フォルダパスもしくはファイル名
Function extractDirOrFile(argFilePath As String, witch As Integer)
    Dim dirs As Variant: dirs = Split(argFilePath, "\")
    If witch = 1 Then
        dirAndFileFromFullPath = Left(argFilePath, Len(argFilePath) - Len(dirs(UBound(dirs))) - 1)
    Else
        dirAndFileFromFullPath = dirs(UBound(dirs))
    End If
End Function
'#説明：指定されたブックに指定されたシートが存在するかどうかを返す
'#引数：wb:走査ブック、checkSheet:探すシート（名/番号どちらも可）
'#戻値：真偽値（true:ある、false:ない）
Function isExistCheckToSheet(wb As Workbook, checkSheet As Variant)
    isExistCheckToSheet = False
    For Each ws In wb.Worksheets
        If Not IsNumeric(checkSheet) Then
            If ws.Name = checkSheet Then
                isExistCheckToSheet = True
            End If
        End If
    Next
    
    If IsNumeric(checkSheet) Then
        '指定値が数値の場合、全シート数より小さい数字かどうかが存在チェックとなる
        isExistCheckToSheet = checkSheet <= wb.Worksheets.count
    End If
End Function
'#説明：指定された値が正規表現パターンに一致するかどうかを検証する
'#引数：text:検査文字列、pattern:正規表現パターン
'#戻値：真偽値（true:一致する、false:一致しない）
Function isRegexpHit(text As String, pattern As String)
    Set regexp = CreateObject("VBScript.RegExp")
    With regexp
         '検索する正規表現条件
         .pattern = pattern
        '大文字小文字の区別（True：しない、False：する）
        .IgnoreCase = False
        '文字列の最後まで検索（True：する、False：しない）
        .Global = True
        '戻り値
        isRegexpHit = .test(text)
    End With
End Function
'#説明：指定された値のうち正規表現パターンに一致するものをコレクションで返す
'#引数：text:検査文字列、pattern:正規表現パターン
'#戻値：コレクション（参照設定：Microsoft VBScript Regular Expressions 5.5）
Function regexpHitCollection(text As String, pattern As String)
    Set regexp = CreateObject("VBScript.RegExp")
    With regexp
         '検索する正規表現条件
         .pattern = pattern
        '大文字小文字の区別（True：しない、False：する）
        .IgnoreCase = False
        '文字列の最後まで検索（True：する、False：しない）
        .Global = True
        '戻り値
        Set regexpHitCollection = .Execute(text)
    End With
End Function
'#説明：指定されたファイルパスがエクセルかどうかを返す
'#引数：filePath:ファイルフルパス
'#戻値：真偽値（true:エクセル、false:エクセル以外）
Function isExcel(filePath As String)
    isExcel = InStr(objFSO.getFile(filePath).Type, "Excel") > 0
End Function
'#説明：指定された文字数分頭から削除した文字列を返す
'#引数：text:文字列、deleteLength:削除する文字数
'#戻値：削除後文字列
Public Function deleteStartText(text As String, Optional deleteLength As Long = 1) As String
    If Len(text) >= deleteLength Then
        deleteStartText = Right(text, Len(text) - deleteLength)
    Else
        deleteStartText = text
    End If
End Function
'#説明：指定された文字数分後ろから削除した文字列を返す
'#引数：text:文字列、deleteLength:削除する文字数
'#戻値：削除後文字列
Public Function deleteEndText(text As String, Optional deleteLength As Long = 1) As String
    If Len(text) >= deleteLength Then
        deleteEndText = Left(text, Len(text) - deleteLength)
    Else
        deleteEndText = text
    End If
End Function
'#説明：指定されたフォルダの最終更新ファイル名を返す
'#引数：argDirPath:フォルダパス
'#戻値：最終更新ファイル名
Function latestFile(argDirPath As String)
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    Dim fileTime As Date
    Dim latestTime As Date
    Dim latestFileName As String
    
    Do While currentFileName <> ""
        'フォルダ＆ファイル名書き出し
        fileTime = FileDateTime(argDirPath & "\" & currentFileName) '取得したファイルの日時を取得
    
        If fileTime > latestTime Then
            '次の比較用
            latestTime = fileTime
            '戻り値用
            latestFileName = currentFileName
        
        End If
        '次のファイル名を取り出す（なければブランク）
        currentFileName = Dir()
    Loop
    
    latestFile = latestFileName
End Function
'#説明：置換配列の内容すべてをテキストに適用する
'#引数：argReplaceArray:置換配列（2次元で1,3列目に置換前・後とする）、argText:置換前の文字列
'#戻値：置換後文字列
Function replaceWithArray(argReplaceArray As Variant, argText As String)

    Dim minIndex As Integer: minIndex = LBound(argReplaceArray, 1)
    Dim replaceText As String: replaceText = argText

    '最初の要素が空の場合、条件指定はなしと判断し、引数をそのまま返す
    If argReplaceArray(minIndex, 1) = "" Then
        replaceWithArray = argText
        Exit Function
    End If
    
    '配列の要素数分走査する
    For i = minIndex To UBound(argReplaceArray)
        replaceText = Replace(replaceText, argReplaceArray(i, 1), argReplaceArray(i, 3))
    Next
    replaceWithArray = replaceText
End Function
'#説明：受け取った内容でテキストファイルを作成する
'#引数：argFilePath:ファイルのフルパス、argContents:ファイルの内容、argCharSet:文字コード
'#戻値：なし
Function createTextFile(argFilePath As String, argContents As String, argCharSet As String)
    With CreateObject("ADODB.Stream")
        .charset = argCharSet
        'https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/lineseparatorsenum?view=sql-server-ver16
        .LineSeparator = 10
        .Open
        .WriteText argContents, 0
        If argCharSet = "UTF-8" Then
            'Streamオブジェクトの先頭からの位置を指定する。Typeに値を設定するときは0である必要がある
            .Position = 0
            '扱うデータ種類をバイナリデータに変更する
            .Type = 1
            '読み取り開始位置？を3バイト目に移動する（3バイトはBOM付き部分を削除するため）
            .Position = 3
            'バイト文字を一時保存
            bytetmp = .read
            'ここでは保存は不要。一度閉じて書き込んだ内容をリセットする目的がある
            .Close
            '再度開いて
            .Open
            'バイト形式で書き込む
            .write bytetmp
        End If
        '保存
        .SaveToFile argFilePath, 2
        'コピー先ファイルを閉じる
        .Close
    End With
End Function
'#説明：テキストファイル内に指定文字列があるかどうか調べる。ある：True、ない：False
'#引数：argFilePath:ファイルのフルパス、searchText:探したい文字列
'#戻値：真偽値（true:ある、false:ない）
Function isInTextFile(argFilePath As String, argSearchArray As Variant)

    'ファイル内容読込
    Dim fileText As String: fileText = getFileText(argFilePath)
    '検索文字列
    Dim searchValue As String
    '見つかった文字列
    Dim findedText As String
    
    For i = LBound(argSearchArray, 1) To UBound(argSearchArray, 1)
        searchValue = argSearchArray(i, 1)
        
        If searchValue <> "" And InStr(fileText, searchValue) > 0 Then
            findedText = findedText & searchValue & ","
        End If
    Next
    isInTextFile = deleteEndText(findedText)
End Function
'#説明：テキストファイル内に指定文字列があるかどうか調べる。
'#引数：argFilePath:ファイルのフルパス、searchText:探したい文字列
'#戻値：真偽値（true:ある、false:ない）
Function isInExcelFile(argFilePath As String, argSearchArray As Variant) 'searchText As String)
    '走査用
    Dim ws As Worksheet
    '検索範囲最初セル
    Dim firstRange As Range
    '検索範囲最終セル
    Dim lastRange As Range
    'ブックを開いて全シートを走査
    Set wb = Workbooks.Open(fileName:=argFilePath, UpdateLinks:=0)
        
    Dim searchValue As String
    Dim findedAddress As String

    For i = 1 To wb.Worksheets.count
        Set ws = wb.Worksheets(i)
        Set firstRange = ws.Range(Split(ws.UsedRange.Address, ":")(0))
        Set lastRange = usedLastRange(ws)

        For j = firstRange.row To lastRange.row
            For k = firstRange.Column To lastRange.Column
                For l = LBound(argSearchArray, 1) To UBound(argSearchArray, 1)
                    searchValue = argSearchArray(l, 1)
                    If searchValue <> "" And InStr(ws.Cells(j, k), searchValue) > 0 Then
                        findedAddress = findedAddress & "[" & ws.Name & ":" & ws.Cells(j, k).Address & ":" & searchValue & "],"
                    End If
                Next
            Next
        Next
    Next
    wb.Close SaveChanges:=False
    isInExcelFile = deleteEndText(findedAddress)
End Function
'#説明：エクセルファイルに置換を施し保存する。
'#引数：argFilePath:ファイルのフルパス、argReplaceArray:置換配列（2次元で1,3列目に置換前・置換後値とする）
'#戻値：置換結果の文字列
Function replaceInExcelFile(argFilePath As String, argReplaceArray As Variant)
    '走査用
    Dim ws As Worksheet
    '検索範囲最初セル
    Dim firstRange As Range
    '検索範囲最終セル
    Dim lastRange As Range
    'ブックを開いて全シートを走査
    Set wb = Workbooks.Open(fileName:=argFilePath, UpdateLinks:=0)

    Dim searchValue As String
    Dim replaceValue As String
    Dim replacedAddress As String
    
    For i = 1 To wb.Worksheets.count
        Set ws = wb.Worksheets(i)
        Set firstRange = ws.Range(Split(ws.UsedRange.Address, ":")(0))
        Set lastRange = usedLastRange(ws)

        For j = firstRange.row To lastRange.row
            For k = firstRange.Column To lastRange.Column
                For l = LBound(argReplaceArray, 1) To UBound(argReplaceArray, 1)
                    searchValue = argReplaceArray(l, 1)
                    replaceValue = argReplaceArray(l, 3)

                    If searchValue <> "" And InStr(ws.Cells(j, k), searchValue) > 0 Then
                        '置換処理
                        ws.Cells(j, k) = Replace(ws.Cells(j, k), searchValue, replaceValue)
                        replacedAddress = replacedAddress & "[" & ws.Name & ":" & ws.Cells(j, k).Address & _
                                                                    ":" & searchValue & ">" & replaceValue & "],"
                        '使い勝手が悪ければ、ここでカスタムログ出力関数を呼ぶようにする
                    End If
                Next
            Next
        Next
    Next
    wb.Close SaveChanges:=True
    replaceInExcelFile = deleteEndText(replacedAddress)
End Function
'#説明：ファイルパスのテキストファイルを作成し、その際に置換処理を行う。
'#引数：argFilePath:ファイルのフルパス、searchText:検索文字列、replaceText:置換文字列
'#戻値：なし
Function replaceInTextFile(argSrcFilePath As String, argDistFilePath As String, argReplaceArray As Variant)
    '文字コード判定（上書き対象と同じ文字コードにするため）
    Dim charset As String: charset = judgeFileCharSet(argSrcFilePath)
    
    Dim searchValue As String
    Dim findedValue As String
    
    If charset = "UTF8" Then
        charset = "UTF-8"
    ElseIf charset = "SJIS" Then
        charset = "SHIFT-JIS"
    End If
    
    With CreateObject("ADODB.Stream")
        .charset = charset
        .Open
        'コピー元ファイルを開く
        .LoadFromFile argSrcFilePath
        'テキスト形式で内容を一時保存
        buf = .ReadText
        
        '文字列置換
        For l = LBound(argReplaceArray, 1) To UBound(argReplaceArray, 1)
            searchValue = argReplaceArray(l, 1)
            If searchValue <> "" And InStr(buf, searchValue) > 0 Then
                buf = Replace(buf, searchValue, argReplaceArray(l, 3))
                findedValue = findedValue & searchValue & ","
            End If
        Next
        
        With CreateObject("ADODB.Stream")
            .charset = charset
            'https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/lineseparatorsenum?view=sql-server-ver16
            .LineSeparator = 10
            .Open
            .WriteText buf, 0
            If charset = "UTF-8" Then
                'Streamオブジェクトの先頭からの位置を指定する。Typeに値を設定するときは0である必要がある
                .Position = 0
                '扱うデータ種類をバイナリデータに変更する
                .Type = 1
                '読み取り開始位置？を3バイト目に移動する（3バイトはBOM付き部分を削除するため）
                .Position = 3
                'バイト文字を一時保存
                bytetmp = .read
                'ここでは保存は不要。一度閉じて書き込んだ内容をリセットする目的がある
                .Close
                '再度開いて
                .Open
                'バイト形式で書き込む
                .write bytetmp
            End If
            '保存
            .SaveToFile argDistFilePath, 2
            'コピー先ファイルを閉じる
            .Close
        End With
        'コピー元ファイルを閉じる
        .Close
    End With
    replaceInTextFile = deleteEndText(findedValue)
End Function
