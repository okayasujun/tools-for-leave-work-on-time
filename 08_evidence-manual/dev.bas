Attribute VB_Name = "dev"
'#作業用のシェイプ整列（開発中にのみ使用する）
'Sub 開発用_シェイプを1か所に整列させる()
'    Dim top As Integer: top = Selection.top
'    Dim left As Integer: left = Selection.left
'
'    For Each moveShape In ActiveSheet.Shapes
'        moveShape.top = top + 20
'        moveShape.left = left + 20
'        top = moveShape.top
'        left = moveShape.left
'    Next
'End Sub
'#セットアップで使用する
'使用中の線の色を調べるプロシージャ。矢印やシェイプを一つだけ選択した状態で実行し、イミディエイトウィンドウを参照
'Sub 開発用_色を調べる()
'    Dim currentColorCode As Long: currentColorCode = Selection.ShapeRange.Item(1).Line.ForeColor.RGB
'    Dim Red As Integer: Red = currentColorCode Mod 256
'    Dim Green As Integer: Green = Int(currentColorCode / 256) Mod 256
'    Dim Blue As Integer: Blue = Int(currentColorCode / 256 / 256)
'
'    Debug.Print "色値：" & currentColorCode
'    Debug.Print "赤：" & Red
'    Debug.Print "緑：" & Green
'    Debug.Print "青：" & Blue
'    Debug.Print "RGB(" & Red & "," & Green; "," & Blue & ")"
'End Sub
'塗りつぶしの色を調べたい場合は､上記ソースの2行目を以下に変更するとOK｡
'    Dim currentColorCode As Long: currentColorCode = Selection.ShapeRange.Item(1).Fill.ForeColor.RGB
'セルの背景色を調べたい場合は､上記ソースの2行目を以下に変更するとOK｡
'    Dim currentColorCode As Long: currentColorCode = Selection.Interior.Color
'セルの文字色を調べたい場合は､上記ソースの2行目を以下に変更するとOK｡
'    Dim currentColorCode As Long: currentColorCode = Selection.Font.Color
'Sub 開発用_すべてのコネクタシェイプを削除する()
'    For Each shp In ActiveSheet.Shapes
'        If shp.Connector Then
'            shp.Delete
'        End If
'    Next
''End Sub
'Sub 選択中シェイプを再現するソースを生成する()
'
'    Dim shp As shape
'    Set shp = Selection.ShapeRange.Item(1)
'    Dim filePath As String: filePath = ActiveWorkbook.Path & "\source-" & Format(Now(), "yyyymmddhhnn") & ".txt"
'    Const CHAR_SET = "SHIFT-JIS"
'    With CreateObject("ADODB.Stream")
'        .Charset = CHAR_SET
'        'https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/lineseparatorsenum?view=sql-server-ver16
'        .LineSeparator = 10
'        .Open
'        .WriteText "    Dim onShape As Object", 1
'        If shp.Type = msoAutoShape Then
'            .WriteText "    Set onShape = ActiveSheet.Shapes.AddShape(" & shp.AutoShapeType & "," & shp.left & "," & shp.top & "," & shp.width & "," & shp.height & ")", 1
'            .WriteText "    onShape.Name = """ & shp.Name & """", 1
'            .WriteText "    onShape.Visible = " & shp.Visible, 1
'            .WriteText "    onShape.Line.ForeColor.RGB = " & shp.Line.ForeColor.RGB, 1
'            '.WriteText "    onShape.ForeColor.RGB = " & shp.ForeColor.RGB, 1
'            '.WriteText "    onShape.TextFrame2.WordArtformat = " & shp.TextFrame2.WordArtformat, 1
'            .WriteText "    onShape.Fill.Transparency = " & shp.Fill.Transparency, 1
'            .WriteText "    onShape.TextFrame.Characters.Text = """ & shp.TextFrame.Characters.Text & """", 1
'            .WriteText "    onShape.Fill.ForeColor.RGB = " & shp.Fill.ForeColor.RGB, 1
'            .WriteText "    onShape.TextFrame2.TextRange.Font.Size = " & shp.TextFrame2.TextRange.Font.Size, 1
'            .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
'            .WriteText "    onShape.TextFrame.Characters.Font.Color = " & shp.TextFrame.Characters.Font.Color, 1
'            .WriteText "    onShape.TextFrame.Characters.Font.Name = """ & shp.TextFrame.Characters.Font.Name & """", 1
'            .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
'            .WriteText "    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = " & shp.TextFrame2.TextRange.ParagraphFormat.Alignment, 1
'            .WriteText "    onShape.Placement = " & shp.Placement, 1
'            .WriteText "    onShape.LockAspectRatio = " & shp.LockAspectRatio, 1
'            .WriteText "    onShape.TextFrame2.AutoSize = " & shp.TextFrame2.AutoSize, 1
'            .WriteText "    onShape.TextFrame2.MarginLeft = " & shp.TextFrame2.MarginLeft, 1
'            .WriteText "    onShape.TextFrame2.MarginRight = " & shp.TextFrame2.MarginRight, 1
'            .WriteText "    onShape.TextFrame2.MarginTop = " & shp.TextFrame2.MarginTop, 1
'            .WriteText "    onShape.TextFrame2.MarginBottom = " & shp.TextFrame2.MarginBottom, 1
'            .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
'            .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
'            .WriteText "    onShape.TextFrame2.HorizontalAnchor = " & shp.TextFrame2.HorizontalAnchor, 1
'            .WriteText "    onShape.TextFrame2.Orientation = " & shp.TextFrame2.Orientation, 1
'
'        ElseIf shp.Connector Then
'            .WriteText "    Set onShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorElbow,BeginX:=0,BeginY:=0,EndX:=0,EndY:=0)", 1
'            .WriteText "    onShape.ConnectorFormat.Type = " & shp.ConnectorFormat.Type, 1
'            .WriteText "    onShape.Name = """ & shp.Name & """", 1
'            .WriteText "    onShape.Line.ForeColor.RGB = " & shp.Line.ForeColor.RGB, 1
'            .WriteText "    onShape.Placement = " & shp.Placement, 1
'            .WriteText "    onShape.LockAspectRatio = " & shp.LockAspectRatio, 1
'            .WriteText "    onShape.top = " & shp.top, 1
'            .WriteText "    onShape.left = " & shp.left, 1
'            .WriteText "    onShape.width = " & shp.width, 1
'            .WriteText "    onShape.height = " & shp.height, 1
'            .WriteText "    onShape.Line.BeginArrowheadStyle = " & shp.Line.BeginArrowheadStyle, 1
'            .WriteText "    onShape.Line.EndArrowheadStyle = " & shp.Line.EndArrowheadStyle, 1
'            .WriteText "    onShape.Line.Weight = " & shp.Line.Weight, 1
'        End If
'        '保存
'        .SaveToFile filePath, 2
'        'コピー先ファイルを閉じる
'        .Close
'    End With
'End Sub
''出力したソースを貼付して動作確認する用
'Sub 出力確認()
'
'End Sub
