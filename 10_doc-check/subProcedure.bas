Attribute VB_Name = "subProcedure"
'###############################
'機能名：ドキュメントチェッカー1
'Author：okayasu jun
'作成日：2022/12/13
'更新日：2022/12/13
'COMMENT：
'###############################

'■実行条件
'シート指定
Dim exeSheet As String
'#チェック処理
Sub check()
    '実行
    Call main
    '完了通知
    MsgBox noticeCount & "件チェックしました。※メッセージは再考"
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
'#機能独自の初期化処理を行う。引数は機能セクションの最初の項目ラベルの場所を指すセル
Function initializeInCustom(customSettingCurrentRange As Range)

    '処理対象のシート情報
    exeSheet = customSettingCurrentRange.Offset(0, 1).Value
    
    'チェック情報の開始セルを取得
    Set customSettingCurrentRange = getBottomEndRange(customSettingCurrentRange, 1)
    
    
    'チェック情報を取得
    Dim exe1StartRange As Range
    Set exe1StartRange = customSettingCurrentRange.Offset(0, 1)
    
    'チェック情報のセル範囲を取得
    Dim exe1ArrayRange As Range
    Dim rightTimes As Integer: rightTimes = Cells(exe1StartRange.Offset(-1, 0).row, Columns.count).End(xlToLeft).Column / 2 - 2
    Set exe1ArrayRange = topSheet.Range(exe1StartRange, regionEndRange(exe1StartRange, headerFlag:=True, rightTimes:=rightTimes))
    
    'セル色条件を取得するための一時処理（色の値設定）
    Dim editRange As Range
    Set editRange = exe1StartRange.Offset(0, 5)
    For i = exe1StartRange.row To getBottomEndRange(exe1StartRange, 1).row
        If topSheet.Cells(i, 9).Interior.Color <> 16777215 Or topSheet.Cells(i, 4) = "セル背景色" Then
            topSheet.Cells(i, 11).Value = topSheet.Cells(i, 9).Interior.Color
        ElseIf topSheet.Cells(i, 4) = "セル文字色（全部）" _
            Or topSheet.Cells(i, 4) = "セル文字色（一部）" _
            Or topSheet.Cells(i, 4) = "ヘッダ文字色" Then
            topSheet.Cells(i, 11).Value = topSheet.Cells(i, 9).Font.Color
        End If
    Next
    
    'チェック情報を配列に格納
    exe1Array = IIf(exe1StartRange.Value = "", exe1StartRange, exe1ArrayRange)
    
    
    'セル色条件を取得するための一時処理（一時値クリア）
'    For i = exe1StartRange.Row To getBottomEndRange(exe1StartRange, 1).Row
'        If topSheet.Cells(i, 11).Value = topSheet.Cells(i, 11).Interior.Color Then
'            topSheet.Cells(i, 11).Value = ""
'        End If
'    Next
    
    '取得地チェック用。まだ使うだろうから一旦残しておく。
'    For i = LBound(exe1Array) To UBound(exe1Array)
'        Debug.Print exe1Array(i, 1) & "：" & exe1Array(i, 4) & "：" & exe1Array(i, 6) & "：" & exe1Array(i, 8)
'    Next
    
    Call initializeLogHeaderSetting
    
End Function
'#ログの設定
Function initializeLogHeaderSetting()
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "フォルダ"
    logSheet.Cells(1, 3) = "ファイル名"
    logSheet.Cells(1, 4) = "シート名"
    logSheet.Cells(1, 5) = "チェック種別"
    logSheet.Cells(1, 6) = "探索行・列番号"
    logSheet.Cells(1, 7) = "期待値"
    logSheet.Cells(1, 8) = "補助情報"
    logSheet.Cells(1, 9) = "結果値"
    logSheet.Cells(1, 10) = "チェック結果"
    logSheet.Cells(1, 11) = "エラー情報"
    logSheet.Cells(1, 12) = "時刻"
    logSheet.Cells(1, 13) = "備考"
End Function
'#対象の全ファイルを走査する。オプションに応じて再帰処理を行う。
Function scanLoopWithFile(argDirPath As String)
    'フォルダ内の最初のファイル名を取得
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '走査中ファイルが条件を通過するかどうか
        If isPassFile(argDirPath, currentFileName) Then
            Call exeCheck(argDirPath, currentFileName)
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
'#
Function exeCheck(argDirPath As String, argFileName As String)
    
    Dim filePath As String: filePath = argDirPath & "\" & argFileName
    
    If isExcel(filePath) Then
        Call scanWithAllSheets(filePath, exeSheet)
    End If
    
End Function

'シートをチェックする
Function customProcess(ws As Worksheet)
    Dim checkType As String
    Dim checkLine As Integer
    Dim expectedValue As String
    Dim assistValue As String

    Dim tempArray As Variant
    Dim tempvalue As Variant
    Dim actualValue As String
    Dim result As String
    Dim errorCount As Integer
    Dim errorLog As String
    Dim headerRowNo As Integer
    
    If IsEmpty(exe1Array) Then
        Exit Function
    End If
    
    Dim latestRow As Integer
    latestRow = ws.Cells(Rows.count, 1).End(xlUp).row
    
    'チェック内容配列を走査する
    For i = LBound(exe1Array) To UBound(exe1Array)
        checkType = exe1Array(i, 1) 'チェック種別
        '探索行/列番号
        checkLine = CInt(exe1Array(i, 4))
        '期待値
        expectedValue = exe1Array(i, 6)
        '補助情報
        assistValue = exe1Array(i, 8)
        result = "○"
        errorLog = ""
        
        'ここ、中身がなかったらcontinueにとばしてもいいな
        
        If checkType = "ヘッダ行取得" Then
            '▼ヘッダ行の場所が正しいかをチェックする（これ、取得できなかったら後続処理やめた方がいいかもな）
            For j = 1 To 20
                If ws.Cells(j, checkLine).Interior.Color = assistValue Then
                    actualValue = j
                    headerRowNo = j
                    Exit For
                End If
            Next
            '結果記録用
            result = IIf(actualValue <> expectedValue, "×", result)
            If headerRowNo = 0 Then
                errorLog = "ヘッダ行が見つからない。"
            End If
            
        ElseIf checkType = "ヘッダ（項目）" Then
            '▼ヘッダの項目名を内容・順番ともに正しいかチェックする
            tempArray = Split(expectedValue, assistValue)
            For j = checkLine To ws.Cells(headerRowNo, Columns.count).End(xlToLeft).Column
                actualValue = actualValue & ws.Cells(headerRowNo, j) & assistValue
            Next
            actualValue = deleteEndText(actualValue)
            '結果記録用
            result = IIf(actualValue <> expectedValue, "×", result)

        ElseIf checkType = "ヘッダ文字色" Then
            '▼ヘッダの文字色を調べる（一部に色がついているものは検知しない）
            For j = checkLine To ws.Cells(headerRowNo, Columns.count).End(xlToLeft).Column
                If ws.Cells(headerRowNo, j).Font.Color <> assistValue Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(headerRowNo, j).Address & "],"
                End If
            Next
            
        ElseIf checkType = "入力規則（リスト）" Then
            '▼検査対象値がリスト内にあるかどうかをチェックする
            tempArray = Split(expectedValue, ",")
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And Not isExistArray(tempArray, ws.Cells(j, checkLine)) Then 'エラーの所属を知らせること
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "入力規則（正規表現）" Then
            '▼検査対象値が正規表現にマッチするかどうかをチェックする
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And Not isRegexpHit(ws.Cells(j, checkLine).Value, expectedValue) Then 'エラーの所属を知らせること
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "入力規則（接頭辞）" Then
            '▼検査対象値が指定地で始まるかどうかをチェックする
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And Not isStartText(ws.Cells(j, checkLine), expectedValue) Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "入力規則（接尾辞）" Then
            '▼検査対象値が指定地で終わるかどうかをチェックする
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And (Not isEndText(ws.Cells(j, checkLine), expectedValue)) Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "入力規則（含文字）" Then
            '▼検査対象値が指定地を含むかどうかをチェックする
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And Not InStr(ws.Cells(j, checkLine), expectedValue) > 0 Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "禁止文字" Then
            '▼検査対象値がリスト内にないかどうかをチェックする
            tempArray = Split(expectedValue, ",")
            For j = headerRowNo + 1 To latestRow
                For k = LBound(tempArray) To UBound(tempArray)
                    tempArray(k) = IIf(tempArray(k) = "改行", vbLf, tempArray(k))
                    If withCheck(exe1Array, CInt(i), ws, CInt(j)) And InStr(ws.Cells(j, checkLine), tempArray(k)) > 0 Then
                        errorCount = errorCount + 1
                        errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                    End If
                Next
            Next
        ElseIf checkType = "重複禁止" Then
            '▼値の重複がないかをチェックする
            For j = headerRowNo + 1 To latestRow
                For k = ws.Cells(Rows.count, 1).End(xlUp).row To j + 1 Step -1
                    If withCheck(exe1Array, CInt(i), ws, CInt(j)) And ws.Cells(j, checkLine) = ws.Cells(k, checkLine) Then
                        errorCount = errorCount + 1
                        errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                    End If
                Next
            Next
            
        ElseIf checkType = "連続性" Then
            '▼値が連続しているかをチェックする
            tempNo = expectedValue
            Dim loopCount As Integer: loopCount = 0
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And tempNo + loopCount <> ws.Cells(j, checkLine) Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
                loopCount = loopCount + 1
            Next
            
        ElseIf checkType = "体裁（目盛線）" Then
            '▼目盛線の設定状態を確認する
            If expectedValue = "非表示" And (Not ActiveWindow.DisplayGridlines) Then
                '期待値・結果値ともに非表示
                actualValue = "非表示"
            ElseIf expectedValue = "表示" And ActiveWindow.DisplayGridlines Then
                '期待値・結果値ともに表示
                actualValue = "表示"
            Else
                errorCount = 1
            End If
            actualValue = IIf(ActiveWindow.DisplayGridlines, "表示", "非表示")
            
        ElseIf checkType = "体裁（縮尺）" Then
            '▼縮尺をチェックする
            If expectedValue <> ActiveWindow.Zoom Then
                errorCount = 1
            End If
            actualValue = ActiveWindow.Zoom
            
        ElseIf checkType = "体裁（ｳｨﾝﾄﾞｳ枠固定）" Then
            '▼ウィンドウ枠の固定が設定されているかどうかをチェックする
            If expectedValue = "未設定" And (Not ActiveWindow.FreezePanes) Then
                '期待値・結果値ともに未設定
                actualValue = "未設定"
            ElseIf expectedValue = "設定" And ActiveWindow.FreezePanes Then
                '期待値・結果値ともに設定
                actualValue = "設定"
            Else
                errorCount = 1
            End If
            actualValue = IIf(ActiveWindow.FreezePanes, "設定", "未設定")
            
        ElseIf checkType = "選択セル位置" Then
            '▼アクティブセルの位置・範囲を調べる
            If expectedValue <> Replace(Selection.Address, "$", "") Then
                errorCount = 1
            End If
            actualValue = Replace(Selection.Address, "$", "")
                
        ElseIf checkType = "使用範囲" Then
            '▼UsedRangeを調べる
            If expectedValue <> Replace(ws.UsedRange.Address, "$", "") Then
                errorCount = 1
            End If
            actualValue = Replace(ws.UsedRange.Address, "$", "")
            
        ElseIf checkType = "シェイプ（数）" Then
            '▼ウィンドウ枠の固定が設定されているかどうかをチェックする
            If expectedValue <> ws.Shapes.count Then
                errorCount = 1
            End If
            actualValue = ws.Shapes.count
        
        ElseIf checkType = "セル背景色" Then
            '▼セル背景色を調べる
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And ws.Cells(j, checkLine).Interior.Color <> assistValue Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & "],"
                End If
            Next

        ElseIf checkType = "セル文字色（全部）" Then
            '▼セルの文字色を調べる（一部に色がついているものは検知しない）
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And ws.Cells(j, checkLine).Font.Color <> assistValue Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & "],"
                End If
            Next
            
        ElseIf checkType = "" Then
        ElseIf checkType = "" Then
        End If
        
        If errorCount > 0 Then
            result = "×"
        End If
        '最後のカンマを削除する
        errorLog = deleteEndText(errorLog)
        
        logSheet.Cells(logWriteLine, 1) = logWriteLine - 1 'No.
        logSheet.Cells(logWriteLine, 2) = ws.Parent.Path 'フォルダ
        logSheet.Cells(logWriteLine, 3) = ws.Parent.Name 'ファイル名
        logSheet.Cells(logWriteLine, 4) = ws.Name 'シート名
        logSheet.Cells(logWriteLine, 5) = checkType 'チェック種別
        logSheet.Cells(logWriteLine, 6) = checkLine '探索行/列番号
        logSheet.Cells(logWriteLine, 7) = expectedValue '期待値
        logSheet.Cells(logWriteLine, 7).WrapText = False
        logSheet.Cells(logWriteLine, 8) = assistValue '補助情報
        logSheet.Cells(logWriteLine, 8).WrapText = False
        logSheet.Cells(logWriteLine, 9) = actualValue '結果値
        logSheet.Cells(logWriteLine, 10) = result 'チェック結果
        logSheet.Cells(logWriteLine, 11) = errorLog 'エラー情報
        logSheet.Cells(logWriteLine, 11).WrapText = False
        time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
        logSheet.Cells(logWriteLine, 12) = time '時刻
        logSheet.Cells(logWriteLine, 13) = errorCount 'エラー数
        logWriteLine = logWriteLine + 1
        '初期化
        actualValue = ""
        result = ""
        errorCount = 0
        
        'ヘッダがなかったら処理を終了する
        If headerRowNo = 0 Then
            Exit Function
        End If
    Next
End Function
'チェック時の付帯条件
'説明：チェック対象かどうかを条件に従い確認する
'引数：checkArray:チェック配列、checkNo:走査中の行番号、ws:チェック対象のシート、row:チェック対象シート内の行番号
'戻値：真偽値（true:チェック対象、false:チェック対象ではない）
Function withCheck(checkArray As Variant, checkNo As Integer, ws As Worksheet, row As Integer)

    Dim checkColumn As Integer
    Dim checkValue As String
    Dim checkType As String
    Dim checkConnector As String
    Dim currentResult As Boolean
    Dim checkedValue As String
    Dim totalResult As Boolean: totalResult = True
    
    Dim checkLogic As String
    Dim checkReturnArray As Variant
    
    Dim collection1 As MatchCollection
    Dim collection2 As MatchCollection
    Dim collection3 As MatchCollection
    Dim collection4 As MatchCollection
    Dim collection5 As MatchCollection
    Dim andor As String
    
    checkLogic = checkArray(checkNo, 9)
    '条件ロジックの有無で評価方法を変える（ロジックか、左から順次か）
    If checkLogic <> "" Then
        '各評価式をy方向の配列に変換する
        checkReturnArray = convertArrayFrom2dTo2d(checkArray, checkNo, 10)
        '条件ロジックから括弧単位のコレクションを取得
        '例：「(1 or 2) and (3 or 4 or 5)」から「(1 or 2)」 と「(3 or 4 or 5)」を抽出）
        Set collection1 = regexpHitCollection(checkLogic, "\(.+?\)")
        '取得できた括弧毎に中の条件を評価する
        For Each hit1 In collection1
            '条件ロジックの数値部分を抽出する
            Set collection2 = regexpHitCollection(hit1.Value, "\d")
            '1 2 3 4 5 を抽出
            For Each hit2 In collection2
                'ループ中の数値から条件列を取得する
                checkColumn = CInt(checkReturnArray(CInt(hit2.Value), 1))
                
                If checkColumn <> 0 Then
                    '検索値
                    checkValue = checkReturnArray(CInt(hit2.Value), 2)
                    '条件種別
                    checkType = checkReturnArray(CInt(hit2.Value), 3)
                    '検査される値
                    checkedValue = ws.Cells(row, checkColumn).Value
                
                    currentResult = getCheckReuslt(checkType, checkedValue, checkValue)
                
                    '括弧内の数値を導出した真偽値で置換
                    checkLogic = Replace(checkLogic, hit2.Value, currentResult)
                End If
            Next
        Next
        '★ここまでで真偽値変換後の文字列が誕生
        Debug.Print "①" & checkLogic
        '括弧毎の論理式を統合する
        Set collection3 = regexpHitCollection(checkLogic, "\(.+?\)")
        For Each hit1 In collection3
            '括弧内の値をコレクションで取得
            Set collection4 = regexpHitCollection(hit1.Value, "[a-zA-Z]{2,5}")
            For i = 0 To collection4.count - 1 Step 2
                If i = 0 Then
                    '最初の条件検証結果はそのまま初期値として設定する
                    totalResult = CBool(collection4.Item(i))
                Else
                    '2つ目以降の条件はAND/ORと前回の結果によって導出する。引数の-1は前回の条件種別を取得したいから
                    currentResult = CBool(collection4.Item(i))
                    andor = CStr(collection4.Item(i - 1))
                    totalResult = deriveTotalResult(totalResult, currentResult, andor)
                End If
            Next
            '括弧を導出した真偽値で置換
            checkLogic = Replace(checkLogic, hit1.Value, totalResult)
        Next
        Debug.Print "②" & checkLogic
        
        '括弧同士の真偽値を
        Set collection5 = regexpHitCollection(checkLogic, "[a-zA-Z]{2,5}")
        For i = 0 To collection5.count - 1 Step 2
            If i = 0 Then
                '最初の条件検証結果はそのまま初期値として設定する
                totalResult = CBool(collection5.Item(i))
            Else
                '2つ目以降の条件はAND/ORと前回の結果によって導出する。引数の-1は前回の条件種別を取得したいから
                currentResult = CBool(collection5.Item(i))
                andor = CStr(collection5.Item(i - 1))
                totalResult = deriveTotalResult(totalResult, currentResult, andor)
            End If
        Next
        Debug.Print "③" & totalResult
    Else

        '10が条件の開始位置
        For i = 10 To UBound(checkArray, 2) Step 8
            '条件列（ブランクなら0になる）
            checkColumn = CInt(checkArray(checkNo, i))
            
            
            If checkColumn <> 0 Then
                '探す値
                checkValue = checkArray(checkNo, i + 2)
                '条件種別
                checkType = checkArray(checkNo, i + 4)
                '論理結合子
                checkConnector = checkArray(checkNo, i - 2)
                '検査される値
                checkedValue = ws.Cells(row, checkColumn).Value
        
                currentResult = getCheckReuslt(checkType, checkedValue, checkValue)
        
                If i = 10 Then
                    '最初の条件検証結果はそのまま初期値として設定する
                    totalResult = currentResult
                Else
                    '現在の最新真偽値と今回のチェック結果を使い論理結合子で導出する
                    totalResult = deriveTotalResult(totalResult, currentResult, checkConnector)
                End If
            End If
        Next
    End If
    withCheck = totalResult
End Function

'
'説明：2次元配列を2次元配列へ
'引数：
'戻値：2次元配列
Function convertArrayFrom2dTo2d(srcArray As Variant, argReadRow As Integer, argReadStartColumn As Integer)
    Dim returnArray() As Variant
    '要素数の設定
    ReDim returnArray(UBound(srcArray, 2), 3)
    Dim count As Integer: count = 1
    
    '二次元目についてループを行う
    For i = argReadStartColumn To UBound(srcArray, 2) Step 8

        If srcArray(argReadRow, i) <> "" Then
            returnArray(count, 0) = count
            '条件列
            returnArray(count, 1) = srcArray(argReadRow, i)
            '検索値
            returnArray(count, 2) = srcArray(argReadRow, i + 2)
            '条件種別
            returnArray(count, 3) = srcArray(argReadRow, i + 4)
            count = count + 1
        End If
    Next
    convertArrayFrom2dTo2d = returnArray
End Function
'説明：
'引数：
'戻値：
Function getCheckReuslt(checkType As String, checkedValue As String, checkValue As String)

    Dim result As Boolean
    
    If checkType = "左記の値と一致する" Then
        result = (checkedValue = checkValue)
    
    ElseIf checkType = "左記の値を含む" Then
        result = InStr(checkedValue, checkValue) > 0
    
    ElseIf checkType = "左記の値を含まない" Then
        result = InStr(checkedValue, checkValue) = 0
    
    ElseIf checkType = "左記の値で始まる" Then
        result = isStartText(checkedValue, checkValue)
    
    ElseIf checkType = "左記の値で終わる" Then
        result = isEndText(checkedValue, checkValue)
    
    End If
    
    getCheckReuslt = result
    
End Function
