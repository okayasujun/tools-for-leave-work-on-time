Attribute VB_Name = "main"
'開始行
Dim fromRow As Integer
'開始列
Dim fromColumn As Integer
'最終行
Dim toRow As Integer
'最終列
Dim toColumn As Integer
'塗りつぶしなしを示す色コード
Const NO_COLOR = 16777215
'黒字を示す文字色コード
Const INIT_TEXT_COLOR = 0
'セルの書式設定の標準
Const BASE_FORMAT = "G/標準"
'状態再現ソースを生成する
Sub createRevivalSource()
    '使用範囲アドレス
    Dim usedRangeAddress As String
    '使用範囲開始アドレス
    Dim fromAddress As String
    '使用範囲最終アドレス
    Dim toAddress As String
    '範囲フラグ（ループの際に2重でやるかどうか。てかこれ使ってねえな）
    Dim doubleLoopFlag As Boolean: doubleLoopFlag = False
    
    '使用範囲取得（値がなく書式だけも対象となる）
    usedRangeAddress = ActiveSheet.UsedRange.Address
    'Debug.Print address
    
    If usedRangeAddress Like "*:*" Then
        '使用範囲情報の取得
        fromAddress = Split(usedRangeAddress, ":")(0)
        toAddress = Split(usedRangeAddress, ":")(1)
        'Debug.Print Range(fromAddress).Row
        'Debug.Print Range(fromAddress).Column
        'Debug.Print Range(toAddress).Row
        'Debug.Print Range(toAddress).Column
        fromRow = Range(fromAddress).Row
        fromColumn = Range(fromAddress).Column
        toRow = Range(toAddress).Row
        toColumn = Range(toAddress).Column
        doubleLoopFlag = True

    Else
        'Debug.Print Range(usedRangeAddress).Row
        'Debug.Print Range(usedRangeAddress).Column
        fromRow = Range(usedRangeAddress).Row
        toRow = Range(usedRangeAddress).Column

    End If
    
    '状態再生します。
    Call writeFromExcelToText
End Sub
'状態再生ソース生成
Function writeFromExcelToText()
    Dim filePath As String ': filePath = ActiveWorkbook.Path & "\setup.bas"
    Dim moduleName As String
    moduleName = InputBox("ファイル名を入れて。拡張子はいらない。", "モジュールファイル名", "setupX")
    filePath = ActiveWorkbook.Path & "\" & moduleName & ".bas"
    
    If StrPtr(moduleName) = 0 Then
        'キャンセル時
        Exit Function
    End If
    
    Const CHAR_SET = "SHIFT-JIS" 'UTF-8 / SHIFT-JIS
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim count As Integer
    Dim functionCount As Integer: functionCount = 1
    Dim temp As String
    

    With CreateObject("ADODB.Stream")
        .Charset = CHAR_SET
        'https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/lineseparatorsenum?view=sql-server-ver16
        .LineSeparator = 10
        .Open
        'ファイル書き出し
        .WriteText "Attribute VB_Name = """ & moduleName & """", 1
        .WriteText "Function revival0()", 1
        .WriteText "    ActiveWindow.DisplayGridlines = " & ActiveWindow.DisplayGridlines, 1
        .WriteText "    ActiveWindow.Zoom = " & ActiveWindow.Zoom, 1
        .WriteText "    ActiveSheet.Name = """ & ActiveSheet.Name & """", 1
        

        For i = fromRow To toRow
            For j = fromColumn To toColumn
                '値
                If ws.Cells(i, j).Value <> "" Then
                    .WriteText "    Cells(" & i & ", " & j & ") = """ & Replace(ws.Cells(i, j), vbLf, """& vbLf & """) & """", 1
                End If
                '書式
                If ws.Cells(i, j).NumberFormatLocal <> BASE_FORMAT Then
                    .WriteText "    Cells(" & i & ", " & j & ").NumberFormatLocal = """ & ws.Cells(i, j).NumberFormatLocal & """", 1
                End If
                '折り返し
                If ws.Cells(i, j).WrapText Then
                    .WriteText "    Cells(" & i & ", " & j & ").WrapText  = " & ws.Cells(i, j).WrapText, 1
                End If
                'フォントサイズ
                If ws.Cells(i, j).Font.Size <> 11 Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Size = " & ws.Cells(i, j).Font.Size, 1
                End If
                'フォント名
                If ws.Cells(i, j).Font.Name <> "" Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Name = """ & ws.Cells(i, j).Font.Name & """", 1
                End If
                '背景色
                If ws.Cells(i, j).Interior.color <> NO_COLOR Then
                    .WriteText "    Cells(" & i & ", " & j & ").Interior.Color = " & ws.Cells(i, j).Interior.color, 1
                End If
                '文字色
                If ws.Cells(i, j).Font.color <> INIT_TEXT_COLOR Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Color = " & ws.Cells(i, j).Font.color, 1
                End If
                '太字
                If ws.Cells(i, j).Font.Bold Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Bold = " & ws.Cells(i, j).Font.Bold, 1
                End If
                'イタリック
                If ws.Cells(i, j).Font.Italic Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Italic = " & ws.Cells(i, j).Font.Italic, 1
                End If
                '取消線
                If ws.Cells(i, j).Font.Strikethrough Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Strikethrough = " & ws.Cells(i, j).Font.Strikethrough, 1
                End If
                '水平位置
                If ws.Cells(i, j).HorizontalAlignment <> xlGeneral Then
                    .WriteText "    Cells(" & i & ", " & j & ").HorizontalAlignment = " & ws.Cells(i, j).HorizontalAlignment, 1
                End If
                '垂直位置
                If ws.Cells(i, j).VerticalAlignment <> xlCenter Then
                    .WriteText "    Cells(" & i & ", " & j & ").VerticalAlignment = " & ws.Cells(i, j).VerticalAlignment, 1
                End If
                'インデントレベル
                If ws.Cells(i, j).IndentLevel > 0 Then
                    .WriteText "    Cells(" & i & ", " & j & ").IndentLevel = " & ws.Cells(i, j).IndentLevel, 1
                End If
                'セルのマージ
                If ws.Cells(i, j).MergeCells Then
                    .WriteText "    Range(""" & ws.Cells(i, j).MergeArea.Item(1).Address(0, 0) & ":" & ws.Cells(i, j).MergeArea.Item(ws.Cells(i, j).MergeArea.count).Address(0, 0) & """).Merge", 1
                End If
                '罫線（上）
                If ws.Cells(i, j).Borders(xlEdgeTop).LineStyle <> xlLineStyleNone Then
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeTop).LineStyle = " & ws.Cells(i, j).Borders(xlEdgeTop).LineStyle, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeTop).color = " & ws.Cells(i, j).Borders(xlEdgeTop).color, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeTop).weight = " & ws.Cells(i, j).Borders(xlEdgeTop).weight, 1
                End If
                '罫線（下）
                If ws.Cells(i, j).Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone Then
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeBottom).LineStyle = " & ws.Cells(i, j).Borders(xlEdgeBottom).LineStyle, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeBottom).color = " & ws.Cells(i, j).Borders(xlEdgeBottom).color, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeBottom).weight = " & ws.Cells(i, j).Borders(xlEdgeBottom).weight, 1
                End If
                '罫線（左）
                If ws.Cells(i, j).Borders(xlEdgeLeft).LineStyle <> xlLineStyleNone Then
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeLeft).LineStyle = " & ws.Cells(i, j).Borders(xlEdgeLeft).LineStyle, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeLeft).color = " & ws.Cells(i, j).Borders(xlEdgeLeft).color, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeLeft).weight = " & ws.Cells(i, j).Borders(xlEdgeLeft).weight, 1
                End If
                '罫線（右）
                If ws.Cells(i, j).Borders(xlEdgeRight).LineStyle <> xlLineStyleNone Then
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeRight).LineStyle = " & ws.Cells(i, j).Borders(xlEdgeRight).LineStyle, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeRight).color = " & ws.Cells(i, j).Borders(xlEdgeRight).color, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeRight).weight = " & ws.Cells(i, j).Borders(xlEdgeRight).weight, 1
                End If
                '入力規則（リスト）
                On Error Resume Next '非設定時のエラー回避
                '変数格納は不要だけど、アクセスすることで設定可否をあぶりだす。それに応じてErrorNumberを評価する
                temp = ws.Cells(i, j).Validation.Type
                If Err.Number = 0 Then
                    .WriteText "    Cells(" & i & ", " & j & ").Validation.Delete", 1
                    .WriteText "    Cells(" & i & ", " & j & ").Validation.Add Type:=xlValidateList, _", 1
                    .WriteText "        Operator:=xlEqual, _", 1
                    .WriteText "        Formula1:=""" & ws.Cells(i, j).Validation.Formula1 & """", 1
                End If
                Err.Clear
                '関数区切り（実行時に起きる「プロシージャが大きすぎます。」を回避するため）
                If count > 30 And count Mod 30 = 0 Then
                    .WriteText "end Function", 1
                    .WriteText "Function revival" & functionCount & "()", 1
                    functionCount = functionCount + 1
                End If
                count = count + 1
            Next
        Next
        For i = fromRow To toRow
            '行幅設定
            .WriteText "    Rows(" & i & ").RowHeight = " & Rows(i).Height, 1
        Next
        For i = fromColumn To toColumn
            '列幅設定
            .WriteText "    Columns(" & i & ").ColumnWidth = " & Columns(i).ColumnWidth, 1
        Next
        
        .WriteText "    Dim onShape As Object", 1
        For Each shp In ActiveSheet.Shapes
        
            If shp.AutoShapeType <> msoShapeMixed Then
                .WriteText "    Set onShape = ActiveSheet.Shapes.AddShape(" & shp.AutoShapeType & "," & shp.Left & "," & shp.Top & "," & shp.Width & "," & shp.Height & ")", 1
                .WriteText "    onShape.Name = """ & shp.Name & """", 1
                .WriteText "    onShape.Visible = " & shp.Visible, 1
                .WriteText "    onShape.Line.ForeColor.RGB = " & shp.Line.ForeColor.RGB, 1
                .WriteText "    onShape.ForeColor.RGB = " & shp.ForeColor.RGB, 1
                '.WriteText "    onShape.TextFrame2.WordArtformat = " & shp.TextFrame2.WordArtformat, 1
                .WriteText "    onShape.Fill.Transparency = " & shp.Fill.Transparency, 1
                .WriteText "    onShape.TextFrame.Characters.Text = """ & shp.TextFrame.Characters.Text & """", 1
                .WriteText "    onShape.Fill.ForeColor.RGB = " & shp.Fill.ForeColor.RGB, 1
                .WriteText "    onShape.TextFrame2.TextRange.Font.Size = " & shp.TextFrame2.TextRange.Font.Size, 1
                .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
                .WriteText "    onShape.TextFrame.Characters.Font.Color = " & shp.TextFrame.Characters.Font.color, 1
                .WriteText "    onShape.TextFrame.Characters.Font.Name = """ & shp.TextFrame.Characters.Font.Name & """", 1
                .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
                .WriteText "    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = " & shp.TextFrame2.TextRange.ParagraphFormat.Alignment, 1
                .WriteText "    onShape.Placement = " & shp.Placement, 1
                .WriteText "    onShape.LockAspectRatio = " & shp.LockAspectRatio, 1
                .WriteText "    onShape.TextFrame2.AutoSize = " & shp.TextFrame2.AutoSize, 1
                .WriteText "    onShape.TextFrame2.MarginLeft = " & shp.TextFrame2.MarginLeft, 1
                .WriteText "    onShape.TextFrame2.MarginRight = " & shp.TextFrame2.MarginRight, 1
                .WriteText "    onShape.TextFrame2.MarginTop = " & shp.TextFrame2.MarginTop, 1
                .WriteText "    onShape.TextFrame2.MarginBottom = " & shp.TextFrame2.MarginBottom, 1
                .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
                .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
                .WriteText "    onShape.TextFrame2.HorizontalAnchor = " & shp.TextFrame2.HorizontalAnchor, 1
                .WriteText "    onShape.TextFrame2.Orientation = " & shp.TextFrame2.Orientation, 1
                
            ElseIf shp.Connector Then
                .WriteText "    Set onShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorElbow,BeginX:=0,BeginY:=0,EndX:=0,EndY:=0)", 1
                .WriteText "    onShape.ConnectorFormat.Type = " & shp.ConnectorFormat.Type, 1
                .WriteText "    onShape.Name = """ & shp.Name & """", 1
                .WriteText "    onShape.Line.ForeColor.RGB = " & shp.Line.ForeColor.RGB, 1
                .WriteText "    onShape.Placement = " & shp.Placement, 1
                .WriteText "    onShape.LockAspectRatio = " & shp.LockAspectRatio, 1
                .WriteText "    onShape.top = " & shp.Top, 1
                .WriteText "    onShape.left = " & shp.Left, 1
                .WriteText "    onShape.width = " & shp.Width, 1
                .WriteText "    onShape.height = " & shp.Height, 1
                .WriteText "    onShape.Line.BeginArrowheadStyle = " & shp.Line.BeginArrowheadStyle, 1
                .WriteText "    onShape.Line.EndArrowheadStyle = " & shp.Line.EndArrowheadStyle, 1
                .WriteText "    onShape.Line.Weight = " & shp.Line.weight, 1

            ElseIf shp.Type = msoFormControl Then
                .WriteText "    Set onShape = ActiveSheet.Buttons.Add(" & shp.Left & "," & shp.Top & "," & shp.Width & "," & shp.Height & ")", 1
                .WriteText "    onShape.OnAction = """ & Mid(shp.OnAction, InStr(shp.OnAction, "!") + 1) & """", 1
                .WriteText "    onShape.Name = """ & shp.Name & """", 1
                .WriteText "    onShape.Visible = " & shp.Visible, 1
                .WriteText "    onShape.Placement = " & shp.Placement, 1
                .WriteText "    onShape.Characters.Text = """"", 1
                'コメントアウト分はなぜか出力されない
                .WriteText "    onShape.Characters.Text = """ & shp.Characters.Text & """", 1
'                .WriteText "    onShape.Text = """ & shp.Text & """", 1
'                .WriteText "    onShape.Caption = """ & shp.Caption & """", 1
'                .WriteText "    onShape.TextFrame2.Characters.Text = """ & shp.TextFrame2.Characters.Text & """", 1
'                .WriteText "    onShape.TextFrame2.TextRange.Font.Size = " & shp.TextFrame2.TextRange.Font.Size, 1
'                .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
'                .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
'                .WriteText "    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = " & shp.TextFrame2.TextRange.ParagraphFormat.Alignment, 1
'                .WriteText "    onShape.TextFrame2.AutoSize = " & shp.TextFrame2.AutoSize, 1
'                .WriteText "    onShape.TextFrame2.MarginLeft = " & shp.TextFrame2.MarginLeft, 1
'                .WriteText "    onShape.TextFrame2.MarginRight = " & shp.TextFrame2.MarginRight, 1
'                .WriteText "    onShape.TextFrame2.MarginTop = " & shp.TextFrame2.MarginTop, 1
'                .WriteText "    onShape.TextFrame2.MarginBottom = " & shp.TextFrame2.MarginBottom, 1
'                .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
'                .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
'                .WriteText "    onShape.TextFrame2.HorizontalAnchor = " & shp.TextFrame2.HorizontalAnchor, 1
'                .WriteText "    onShape.TextFrame2.Orientation = " & shp.TextFrame2.Orientation, 1
            
            End If
            '関数区切り（実行時に起きる「プロシージャが大きすぎます。」を回避するため）
            If count > 30 And count Mod 30 = 0 Then
                .WriteText "end Function", 1
                .WriteText "Function revival" & functionCount & "()", 1
                functionCount = functionCount + 1
            End If
            count = count + 1
                
        Next
        '最後の1行は改行なし
        .WriteText "End Function", 1
        
        .WriteText "Sub revival()", 1
        If IsNumeric(Right(moduleName, 1)) Then
            'モジュール名の最後が数値の場合、シート作成処理を追加する
            .WriteText "    Worksheets.Add After:=Worksheets(Worksheets.Count)", 1
        End If
        
        For i = 0 To functionCount - 1
            .WriteText "    CALL revival" & i & "()", 1
        Next
        
        If IsNumeric(Right(moduleName, 1)) Then
            .WriteText "    Worksheets(1).select", 1
        End If
        .WriteText "end sub", 0
        
        If CHAR_SET = "UTF-8" Then
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
        .SaveToFile filePath, 2
        'コピー先ファイルを閉じる
        .Close
    End With
End Function
