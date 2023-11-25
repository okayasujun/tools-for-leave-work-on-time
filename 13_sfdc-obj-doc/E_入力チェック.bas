Attribute VB_Name = "E_入力チェック"
Private errorSheet As Worksheet
Private logSheet As Worksheet
Private lastRow As Integer
Private lastColumn As Integer
Private dataType As String
Private errorCheckColumn As Integer
Private errorMsgFormat As String
Private errorLastRow As Integer
Private condition As String
Private logWriteLine As Integer
Private i As Integer
Private j As Integer
Sub G_入力チェック()
    '▼TODO
    '・API名のルール不履行がないか
    '・API名に重複はないか
    '・オブジェクトの設定と整合しているか
    '・API名の頭文字は大文字であれ（これはいらんな）
    '・オブジェクトの共有設定と主従関係項目の有無
    Call initiarize
    Call init
    Call repaint
    
    For i = 5 To lastRow
        With itemSheet
            dataType = .Cells(i, 7).Value
            dataType = IIf(.Cells(i, 8).Value = ON_TRUE, "(数式)" & dataType, dataType)
            
            For j = 2 To errorSheet.Cells(1, 1).End(xlDown).row
                condition = errorSheet.Cells(j, 6).Value
                errorCheckColumn = errorSheet.Cells(j, 4).Value
                errorMsgFormat = errorSheet.Cells(j, 7).Value
            
                '有効項目で、データ型が一致すればチェックにかける
                If .Cells(i, 2).Value = ON_TRUE And errorSheet.Cells(j, 2).Value = dataType Then
                
                    If condition = "以上" Then
                        If Not .Cells(i, errorCheckColumn).Value >= errorSheet.Cells(j, 5).Value Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    ElseIf condition = "以下" Then
                        If Not .Cells(i, errorCheckColumn).Value <= errorSheet.Cells(j, 5).Value Then
                            Call setError(errorMsgFormat)
                        End If
                    
                    ElseIf condition = "等しい" Then
                        If Not .Cells(i, errorCheckColumn).Value = errorSheet.Cells(j, 5).Value Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    ElseIf condition = "必須" Then
                        If Not .Cells(i, errorCheckColumn).Value <> "" Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    ElseIf condition = "正規表現に一致する" And dataType = "選択リスト" Then
                        Call setupRegexp(errorSheet.Cells(j, 5))
                        If Not UBound(Split(.Cells(i, errorCheckColumn).Value, vbLf)) + 1 = regexp.Execute(.Cells(i, errorCheckColumn).Value).Count Then
                            Call setError(errorMsgFormat)
                            logSheet.Cells(logWriteLine - 1, 2).WrapText = False
                        End If
                        
                    ElseIf condition = "と一致する" Then
                        Dim targetArray() As String: targetArray = Split(errorSheet.Cells(j, 5).Value, ",")
                        If Not isExistArray(targetArray, .Cells(i, errorCheckColumn).Value) Then
                            Call setError(errorMsgFormat)
                        End If
                    End If
                    
                '[Error定義]上、データ型未指定のケース
                ElseIf .Cells(i, 2).Value = "〇" And errorSheet.Cells(j, 2).Value = "-" Then
                    If errorSheet.Cells(j, 5) = "改行文字" And condition = "含まない" Then
                        If .Cells(i, errorCheckColumn).Value Like "*" & vbLf & "*" _
                            Or .Cells(i, errorCheckColumn).Value Like "*" & vbCrLf & "*" _
                            Or .Cells(i, errorCheckColumn).Value Like "*" & vbCr & "*" Then
                            Call setError(errorMsgFormat)
                            logSheet.Cells(logWriteLine - 1, 2).WrapText = False
                        End If
                        
                    ElseIf condition = "必須" Then
                        If Not .Cells(i, errorCheckColumn).Value <> "" Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    ElseIf condition = "字以下" Then
                        If Not Len(.Cells(i, errorCheckColumn).Value) <= errorSheet.Cells(j, 5).Value Then
                            Call setError(errorMsgFormat)
                        End If
                    
                    ElseIf condition = "と一致する" Then
                        targetArray = Split(errorSheet.Cells(j, 5).Value, ",")
                        If Not isExistArray(targetArray, .Cells(i, errorCheckColumn).Value) Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    End If
                End If
            Next
        End With
    Next
    
    If logWriteLine = 2 Then
        MsgBox "入力チェックが完了しました。" & vbCrLf & "入力不備はありません。"
    Else
        MsgBox "入力不備があります。[log]シートを確認してください。"
        logSheet.Select
    End If
End Sub
'エラー情報を設定する
Private Function setError(format As String)
    logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
    logSheet.Cells(logWriteLine, 2) = getMessage(format, _
                                i, itemSheet.Cells(i, 3), errorSheet.Cells(j, 3), _
                                errorSheet.Cells(j, 5), errorSheet.Cells(j, 6), _
                                itemSheet.Cells(i, errorCheckColumn).Value)
    logWriteLine = logWriteLine + 1
    '対象セルハイライト
    itemSheet.Cells(i, errorCheckColumn).Interior.Color = RGB(255, 255, 0)
End Function
'チェック処理用初期化処理
Private Function init()
    Set errorSheet = Sheets("Error定義")
    Set logSheet = Sheets("log")
    lastRow = itemSheet.Cells(5, 1).End(xlDown).row
    lastColumn = itemSheet.Cells(4, 1).End(xlToRight).column
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 1).Font.Bold = True
    logSheet.Cells(1, 2) = "内容"
    logSheet.Cells(1, 2).Font.Bold = True
    logWriteLine = 2
End Function
'セル塗色を適切な状態にする
Function repaint()

    For i = 5 To lastRow
        With itemSheet
            dataType = .Cells(i, 7).Value
            dataType = IIf(.Cells(i, 8).Value = "〇", "(数式)" & dataType, dataType)
            
            If .Cells(i, 2).Value = "×" Then
                'すべてグレーに塗る
                .Range(.Cells(i, 1), .Cells(i, 38)).Interior.Color = RGB(191, 191, 191)
            Else
                '色分け
                '走査中行のデータ型を定義シートから探す
                For j = 4 To 31
                    If dataType = itemMetaSheet.Cells(2, j).Value Then
                        errorCheckColumn = j
                        Exit For
                    End If
                Next
                    
                'データ型に応じて記載不要な情報をグレーに塗色する
                For j = 3 To 37
                    colorChangeColumn = itemMetaSheet.Cells(j, 2).Value
                    If itemMetaSheet.Cells(j, errorCheckColumn).Value And colorChangeColumn > 0 Then
                        '有効セル
                        .Cells(i, colorChangeColumn).Interior.Color = NO_COLOR
                    ElseIf colorChangeColumn > 0 Then
                        '無効セル
                        .Cells(i, colorChangeColumn).Interior.Color = RGB(191, 191, 191)
                        .Cells(i, colorChangeColumn).Value = ""
                    End If
                Next
                '上記処理で掬えてない分
                .Cells(i, 1).Interior.Color = NO_COLOR
                .Cells(i, 2).Interior.Color = NO_COLOR
                .Cells(i, 4).Interior.Color = NO_COLOR
                .Cells(i, 6).Interior.Color = NO_COLOR
                .Cells(i, 7).Interior.Color = NO_COLOR
                .Cells(i, 8).Interior.Color = NO_COLOR
            End If
        End With
    Next

End Function

