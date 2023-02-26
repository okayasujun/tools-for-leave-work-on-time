Attribute VB_Name = "main"
Sub createFlow()
    Const CENTER = 450
    '■0.基準位置指定、シェイプ表取得
    
    '■1.シェイプを作成する
    '名前、出現（形のパートはどこか別に用意しよう）
    '■2.場所移動
    '■3.コネクタ付与
    '
    '
    '
    '
    '
    '
    '
    Dim onShape As Shape
    Dim hjShape As Shape
    Dim shapeList As Variant: shapeList = Sheets("シェイプ表２").Range("A2:F31")
    Dim shapeType As Integer
    Dim lastRow As Integer: lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Dim cellColorCode As Long
    'Dim baseCenterPoint As Integer: baseCenterPoint = CENTER
    '■1.シェイプ作成
    For i = 8 To lastRow
        'フロー作成
        shapeType = vlookup(shapeList, Cells(i, 3), 2, 4)
        cellColorCode = Cells(i, 2).Interior.Color
        Set onShape = ActiveSheet.Shapes.AddShape(shapeType, 400, 100, 100, 30)
        onShape.TextFrame.Characters.Text = Cells(i, 4)
        onShape.Name = Cells(i, 2)
        Set onShape = baseShape(onShape)
        onShape.Fill.ForeColor.RGB = RGB(cellColorCode Mod 256, Int(cellColorCode / 256) Mod 256, Int(cellColorCode / 256 / 256))
        
        'カスタムがあればここに分岐をかく
        If Cells(i, 3) = "ループ開始" Then
            onShape.AutoShapeType = 156
            onShape.Adjustments.Item(1) = 0.3
            onShape.Adjustments.Item(2) = 0
        ElseIf Cells(i, 3) = "ループ終了" Then
            onShape.AutoShapeType = 156
            onShape.Adjustments.Item(1) = 0
            onShape.Adjustments.Item(2) = 0.3
        ElseIf Cells(i, 3) = "参照" Then
            onShape.height = 30
            onShape.width = 30
        End If
    Next
    
    '■2.場所移動
    Dim moveShape As Shape
    Dim baseTopPoint As Integer: baseTopPoint = 100
    Dim top As Integer: top = 100
    Dim nextShape As String
    Dim nextShapeArray As Variant
    Dim nextNo As String
    Dim movedShape As String
    For i = 7 To lastRow
        If i = 7 Then
            'ここダサいな。再考しよう。
            nextShape = CStr(1)
        Else
            nextShape = CStr(Cells(i, 11))
        End If
        
        If nextShape Like "*" & vbLf & "*" Then
            '複数ある場合
            nextShapeArray = Split(nextShape, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
                '遷移先のグループ（ここ横並びのmoveもやっちゃうか）
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                Set moveShape = ActiveSheet.Shapes(nextNo)
                '中央揃え
                moveShape.Left = CENTER - moveShape.width / 2
                moveShape.top = top
                movedShape = movedShape & nextNo & ","
                
                moveShape.Left = CENTER - (UBound(nextShapeArray) - j) * moveShape.width * 1.3
                '■ここで小さなシェイプ生成やってみるか
                Set hjShape = ActiveSheet.Shapes.AddShape(shapeType, 40, 10, 10, 10)
                hjShape.TextFrame.Characters.Text = Split(CStr(nextShapeArray(j)), ":")(0)
                Set hjShape = hojoShape(hjShape)
                hjShape.Left = moveShape.Left - 20
                hjShape.top = moveShape.top - 20
            Next
            top = top + moveShape.height + 25
        ElseIf Not movedShape Like "*" & nextShape & "*" Then
            '単一でかつ、未処理だったときだけ行う
            Set moveShape = ActiveSheet.Shapes(nextShape)
            '中央揃え
            moveShape.Left = CENTER - moveShape.width / 2
            moveShape.top = top
            movedShape = movedShape & nextShape & ","
            top = top + moveShape.height + 25
        End If
    Next
    
    '■3.コネクタ付与
    Dim beforeShape As Shape
    Dim afterShape As Shape
    Dim beforeShapeName As String
    Dim afterShapeName As String
    Dim connectShape As Shape
    Dim startPoint As Integer
    Dim endPoint As Integer
    Dim elbowNo As String
    For i = 8 To lastRow
        'beforeシェイプ
        beforeShapeName = CStr(Cells(i, 2))
        Set beforeShape = ActiveSheet.Shapes(beforeShapeName)
        startPoint = vlookup(shapeList, beforeShape.AutoShapeType, 4, 5)
        
        
        'afterシェイプ
        afterShapeName = CStr(Cells(i, 11))
        If afterShapeName Like "*" & vbLf & "*" Then
            nextShapeArray = Split(afterShapeName, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
                'ここかもしくは
                Set connectShape = setConnect(connectShape)
                
                connectShape.ConnectorFormat.BeginConnect beforeShape, startPoint
                '遷移先のグループ（ここ横並びのmoveもやっちゃうか）
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                'もしもbeforeシェイプが分岐系ならエルボーコネクタ
                If beforeShape.AutoShapeType = 63 Then
                    connectShape.ConnectorFormat.Type = msoConnectorElbow
                    elbowNo = elbowNo & nextNo & ","
                End If
                Set afterShape = ActiveSheet.Shapes(nextNo)
                endPoint = vlookup(shapeList, afterShape.AutoShapeType, 4, 6)
                connectShape.ConnectorFormat.EndConnect afterShape, endPoint
            Next
        ElseIf i <> lastRow Then
            Set afterShape = ActiveSheet.Shapes(afterShapeName)
            
            Set connectShape = setConnect(connectShape)
            connectShape.ConnectorFormat.BeginConnect beforeShape, startPoint
            endPoint = vlookup(shapeList, afterShape.AutoShapeType, 4, 6)
            connectShape.ConnectorFormat.EndConnect afterShape, endPoint
            'こっちにいる
            If elbowNo Like "*" & beforeShapeName & "," & "*" Then
                connectShape.ConnectorFormat.Type = msoConnectorElbow
            End If
        End If
    
    Next
    
End Sub
Function setConnect(connectShape As Shape)
    'コネクトシェイプ
    Set connectShape = ActiveSheet.Shapes.AddConnector( _
        Type:=1, _
        BeginX:=10, _
        BeginY:=1000, _
        EndX:=10, _
        EndY:=90)
    '■終点コネクタを三角に。
    connectShape.Line.EndArrowheadStyle = msoArrowheadTriangle
    '■線の色
    connectShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    '■線の太さ
    connectShape.Line.Weight = 1
    Set setConnect = connectShape
End Function
'表から使用したいシェイプを探す
Function vlookup(list As Variant, searchVal As String, searchCol As Integer, returnCol As Integer)
    vlookup = "61"
    For i = LBound(list) To UBound(list)
        If list(i, searchCol) Like "*" & searchVal & "*" Then
            vlookup = list(i, returnCol)
            Exit For
        End If
    Next
    
End Function
'コネクトの種類を直線とエルボーで切り替える
Sub connectChange()
    If Selection.ShapeRange.ConnectorFormat.Type = msoConnectorElbow Then
        Selection.ShapeRange.ConnectorFormat.Type = msoConnectorStraight
    Else
        Selection.ShapeRange.ConnectorFormat.Type = msoConnectorElbow
    End If
End Sub
'複数選択したシェイプのうち、最初のシェイプの左座標に他のシェイプの位置を合わせる
Sub leftPointSameSet()
    Dim baseLeftPoint As Integer: baseLeftPoint = Selection.ShapeRange.Item(1).Left
    For Each sp In Selection.ShapeRange
        sp.Left = baseLeftPoint
    Next
End Sub
'選択したコネクタの始点側のシェイプとの接続位置を変更する
Sub checkExistsConnector()
    Dim currentBeginConnectPoint As Integer
    Dim targetShape As Shape
    
    '現在の接続ポイント取得
    currentBeginConnectPoint = Selection.ShapeRange.ConnectorFormat.BeginConnectionSite
    '始点側のシェイプ取得
    Set targetShape = Selection.ShapeRange.ConnectorFormat.BeginConnectedShape
    '次の接続ポイントに変更
    If currentBeginConnectPoint = targetShape.ConnectionSiteCount Then
        currentBeginConnectPoint = 1
    Else
        currentBeginConnectPoint = currentBeginConnectPoint + 1
    End If
    '接続ポイントを変更
    Selection.ShapeRange.ConnectorFormat.BeginConnect targetShape, currentBeginConnectPoint
End Sub
Function baseShape(onShape As Shape)
    '■塗りつぶし（msoTrue:あり、msoFalse:なし）
    onShape.Fill.Visible = msoTrue
    '■線の太さ。お好みでどうぞ
    onShape.Line.Weight = 1
    '■色指定
    onShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    '■塗りつぶし色指定
    onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    '■文字色
    onShape.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    '■書体
    onShape.TextFrame.Characters.Font.Name = "Meiryo UI"
    '文字垂直寄せ
    onShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    '文字水平寄せ
    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        'テキスト設定後にしないといけない
        onShape.TextFrame2.WordWrap = msoFalse
        onShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        If onShape.width < 100 Then
            onShape.TextFrame2.AutoSize = msoAutoSizeNone
            onShape.width = 100
        End If
        onShape.height = 30
    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    Set baseShape = onShape
End Function
Function hojoShape(onShape As Shape)
    '■塗りつぶし（msoTrue:あり、msoFalse:なし）
    onShape.Fill.Visible = msoFalse
    '■線の太さ。お好みでどうぞ
    onShape.Line.Weight = 0
    onShape.Line.Visible = msoFalse
    '■色指定
    onShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    '■塗りつぶし色指定
    'onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    '■文字色
    onShape.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    '■書体
    onShape.TextFrame.Characters.Font.Name = "Meiryo UI"
    '文字垂直寄せ
    onShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    '文字水平寄せ
    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        'テキスト設定後にしないといけない
        onShape.TextFrame2.WordWrap = msoFalse
        onShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        If onShape.width < 30 Then
            onShape.TextFrame2.AutoSize = msoAutoSizeNone
            onShape.width = 30
        End If
        onShape.height = 15
    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    Set hojoShape = onShape
End Function
