Attribute VB_Name = "G_入力チェック"
Private itemSheet As Worksheet
Private itemMetaSheet As Worksheet
Private errorSheet As Worksheet
Private logSheet As Worksheet
Private lastRow As Integer
Private lastColumn As Integer
Private dataType As String
Private dataTypeColumn As Integer
Private errorLastRow As Integer
Private condition As String
Private logWriteLine As Integer
Const NO_COLOR = 16777215
Sub G_入力チェック()
    '・API名のルール不履行がないか
    '・API名に重複はないか
    '・オブジェクトの設定と整合しているか
    '・API名の頭文字は大文字であれ（これはいらんな）
    '・オブジェクトの共有設定と主従関係項目の有無
    Call init
    Call repaint
    
    For i = 5 To lastRow
        With itemSheet
            dataType = .Cells(i, 7).Value
            dataType = IIf(.Cells(i, 8).Value = "〇", "(数式)" & dataType, dataType)
            
            For j = 2 To errorSheet.Cells(1, 1).End(xlDown).row
                    condition = errorSheet.Cells(j, 6).Value
                    dataTypeColumn = errorSheet.Cells(j, 4).Value
            
                If .Cells(i, 2).Value = "〇" And errorSheet.Cells(j, 2).Value = dataType Then
                
                    If condition = "以上" Then
                        If Not .Cells(i, dataTypeColumn).Value >= errorSheet.Cells(j, 5).Value Then
                            '制約エラー
                            logSheet.Cells(logWriteLine, 1) = i & "行目の「" & .Cells(i, 3) & "」項目は「" _
                                    & errorSheet.Cells(j, 3) & "」を" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "にしてください。(" & .Cells(i, dataTypeColumn).Value & ")"
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
                        End If
                        
                    ElseIf condition = "以下" Then
                        If Not .Cells(i, dataTypeColumn).Value <= errorSheet.Cells(j, 5).Value Then
                            '制約エラー
                            logSheet.Cells(logWriteLine, 1) = i & "行目の「" & .Cells(i, 3) & "」項目は「" _
                                    & errorSheet.Cells(j, 3) & "」を" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "にしてください。(" & .Cells(i, dataTypeColumn).Value & ")"
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
                        End If
                    
                    ElseIf condition = "等しい" Then
                        If Not .Cells(i, dataTypeColumn).Value = errorSheet.Cells(j, 5).Value Then
                            logSheet.Cells(logWriteLine, 1) = i & "行目の「" & .Cells(i, 3) & "」項目は「" _
                                    & errorSheet.Cells(j, 3) & "」を" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "にしてください。(" & .Cells(i, dataTypeColumn).Value & ")"
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
                        End If
                        
                    ElseIf condition = "必須" Then
                        If Not .Cells(i, dataTypeColumn).Value <> "" Then
                            '制約エラー
                            logSheet.Cells(logWriteLine, 1) = i & "行目の「" & .Cells(i, 3) & "」項目は「" _
                                    & errorSheet.Cells(j, 3) & "」を" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "にしてください。(" & .Cells(i, dataTypeColumn).Value & ")"
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
                        End If
                    End If
                ElseIf .Cells(i, 2).Value = "〇" And errorSheet.Cells(j, 2).Value = "-" Then
                    If errorSheet.Cells(j, 5) = "改行文字" And condition = "含まない" Then
                        If .Cells(i, dataTypeColumn).Value Like "*" & vbLf & "*" _
                        Or .Cells(i, dataTypeColumn).Value Like "*" & vbCrLf & "*" _
                        Or .Cells(i, dataTypeColumn).Value Like "*" & vbCr & "*" Then
                            
                            logSheet.Cells(logWriteLine, 1) = i & "行目の「" & .Cells(i, 3) & "」項目は「" _
                                    & errorSheet.Cells(j, 3) & "」を" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "にしてください。(" & .Cells(i, dataTypeColumn).Value & ")"
                            
                            logSheet.Cells(logWriteLine, 1).WrapText = False
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
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
    End If
    
End Sub
Function init()
    Set itemSheet = Sheets(ITEM_SHEET)
    Set itemMetaSheet = Sheets(ITEM_META_SHEET)
    Set errorSheet = Sheets("Error定義")
    Set logSheet = Sheets("log")
    lastRow = itemSheet.Cells(5, 1).End(xlDown).row
    lastColumn = itemSheet.Cells(4, 1).End(xlToRight).column
    logSheet.Cells.Clear
    logWriteLine = 2
End Function
Function repaint()

    For i = 5 To lastRow
        With itemSheet
            dataType = .Cells(i, 7).Value
            dataType = IIf(.Cells(i, 8).Value = "〇", "(数式)" & dataType, dataType)
            
            If .Cells(i, 2).Value = "×" Then
                'すべてグレーに塗る
                .Range(Cells(i, 1), Cells(i, 38)).Interior.Color = RGB(191, 191, 191)
            Else
                '色分け
                '走査中行のデータ型を定義シートから探す
                For j = 4 To 31
                    If dataType = itemMetaSheet.Cells(2, j).Value Then
                        dataTypeColumn = j
                        Exit For
                    End If
                Next
                    
                'データ型に応じて記載不要な情報をグレーに塗色する
                For j = 3 To 37
                    colorChangeColumn = itemMetaSheet.Cells(j, 2).Value
                    If itemMetaSheet.Cells(j, dataTypeColumn).Value = "〇" And colorChangeColumn > 0 Then
                        .Cells(i, colorChangeColumn).Interior.Color = NO_COLOR
                    ElseIf colorChangeColumn > 0 Then
                        .Cells(i, colorChangeColumn).Interior.Color = RGB(191, 191, 191)
                        .Cells(i, colorChangeColumn).Value = ""
                    End If
                Next
            End If
        End With
    Next

End Function
Sub a()
    Debug.Print Selection.Interior.Color
End Sub
Sub currentColorCheck()
    Dim currentColorCode As Long: currentColorCode = Selection.Interior.Color
    Dim Red As Integer: Red = currentColorCode Mod 256
    Dim Green As Integer: Green = Int(currentColorCode / 256) Mod 256
    Dim Blue As Integer: Blue = Int(currentColorCode / 256 / 256)
    
    Debug.Print "色値：" & currentColorCode
    Debug.Print "赤：" & Red
    Debug.Print "緑：" & Green
    Debug.Print "青：" & Blue
    Debug.Print "RGB(" & Red & "," & Green; "," & Blue & ")"
End Sub
