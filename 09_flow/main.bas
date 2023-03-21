Attribute VB_Name = "main"
Const PROCESS_NO_COL = 1
Const FLOW_NO_COL = 2
Const AFTER_NO_COL = 3
Const TYPE_COL = 4
Const FLOW_TEXT_COL = 5

Const LOOP_START_ROW = 2
Const HEIGHT_MARGIN = 32
Const WIDTH_MARGIN = 30
Dim LAST_FLOW_NO As String
Dim center As Integer
Dim shapeList As Variant
Dim flowList As Variant
'作成シェイプ
Dim onShape As Shape
   
'標準シェイプ
Dim hjShape As Shape
Function adjust()
    Dim currentShape As String
    Dim nextShape As String
    Dim nextShapeArray As Variant
    Dim beforeShape As Shape
    Dim moveShape As Shape
    Dim switchTotalWidth As Integer
    Dim TOP_CENTER As Integer: TOP_CENTER = ActiveSheet.Shapes("1").Left + ActiveSheet.Shapes("1").Width / 2
    'シェイプのタイプを表す数字
    Dim shapeType As String
    For i = LBound(flowList) To UBound(flowList)
        shapeType = CStr(flowList(i, TYPE_COL))
        '遷移先情報の取得
        currentShape = CStr(flowList(i, FLOW_NO_COL))
        Set beforeShape = ActiveSheet.Shapes(currentShape)
        nextShape = CStr(flowList(i, AFTER_NO_COL))
        
        If nextShape Like "*" & vbLf & "*" Then
            '遷移先が複数ある場合
            nextShapeArray = Split(nextShape, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
            
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                
                Set moveShape = ActiveSheet.Shapes(nextNo)
                
                If UBound(nextShapeArray) = 1 And Not shapeType Like "ループ*" Then
                    If j = 0 Then
                        moveShape.top = beforeShape.top
                        moveShape.Left = beforeShape.Left - moveShape.Width - WIDTH_MARGIN
                        ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.BeginConnect beforeShape, 2
                        ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.EndConnect moveShape, 4
                        ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.Type = msoConnectorStraight
                        
'                        ActiveSheet.Shapes(moveShape.Name & "の補助").top = ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).top + 2
'                        ActiveSheet.Shapes(moveShape.Name & "の補助").Left = ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).Left + 2
                    ElseIf j = 1 Then
                        '前シェイプとセンターを合わせる
                        Call setCenterPosition(moveShape, beforeShape.Left + beforeShape.Width / 2)
                        ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.Type = msoConnectorStraight
'                        ActiveSheet.Shapes(moveShape.Name & "の補助").top = ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).top + 2
'                        ActiveSheet.Shapes(moveShape.Name & "の補助").Left = ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).Left + 2
                    End If
                
                ElseIf UBound(nextShapeArray) > 1 Then
                    For k = LBound(nextShapeArray) To UBound(nextShapeArray)
                        nextNo = Split(CStr(nextShapeArray(k)), ":")(1)
                        switchTotalWidth = switchTotalWidth + ActiveSheet.Shapes(nextNo).Width + WIDTH_MARGIN
                    Next
                    a = switchTotalWidth / (UBound(nextShapeArray)) * (j)
                    'moveShape.Left = a
                    moveShape.Left = TOP_CENTER - switchTotalWidth / 2 + a - (moveShape.Width / 2)
                    switchTotalWidth = 0
                        ActiveSheet.Shapes(moveShape.Name & "の補助").top = ActiveSheet.Shapes(moveShape.Name).top - 16
                        ActiveSheet.Shapes(moveShape.Name & "の補助").Left = ActiveSheet.Shapes(moveShape.Name).Left + ActiveSheet.Shapes(moveShape.Name).Width / 2
                End If
            
                '遷移先のグループ
            Next
        Else
            '遷移先が単一
            Set moveShape = ActiveSheet.Shapes(nextShape)
            If nextShape = LAST_FLOW_NO Then
                '接続始点を変える
                ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.BeginConnect beforeShape, 2
                    ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.EndConnect moveShape, 2
            End If
        End If
    Next
    'この辺で終了要素のコネクタを最適化できるか？
End Function
Function init()

    'センターポジション。定数扱いなので大文字定義にする
     center = Selection.Left + Selection.Width / 2
    
    '使用するシェイプのリスト
    shapeList = Sheets("シェイプ一覧").Range("A2:F31")
    
    'シェイプのタイプを表す数字
    Dim shapeType As Integer
    
    '最終行（フローNoで取得する）
    Dim lastRow As Integer: lastRow = Sheets("詳細").Cells(Rows.Count, 2).End(xlUp).Row
    
    'フローの最終番号
    LAST_FLOW_NO = Sheets("詳細").Cells(lastRow, 2)
    
    '使用するシェイプのリスト
    flowList = Sheets("詳細").Range("A2:R" & lastRow)
End Function
Sub A_フロー作成()
Attribute A_フロー作成.VB_ProcData.VB_Invoke_Func = "q\n14"

    Call init

    '■1.シェイプ作成
    Call createFlowParts
    
    '■2.場所移動
    Call moveFlowParts

    '■3.コネクタ付与
    Call addConnector
    
    '■4.調整
    Call adjust

End Sub
Function createFlowParts()
    '走査対象セルの色コード
    Dim cellColorCode As Long
    '■1.シェイプ作成
    For i = LBound(flowList) To UBound(flowList)
        'シェイプ種別取得
        shapeType = vlookup(shapeList, flowList(i, TYPE_COL), 2, 4)
        'フローNo列のセル色を取得
        cellColorCode = 16777215 'flowList(i, 2).Interior.Color
        'シェイプを生成
        Set onShape = ActiveSheet.Shapes.AddShape(shapeType, 400, 100, 100, 30)
        'テキストの設定
        onShape.TextFrame.Characters.Text = flowList(i, FLOW_TEXT_COL)
        '名前の設定
        onShape.Name = flowList(i, 2)
        'シェイプ生成
        Set onShape = baseShape(onShape)
        'シェイプの色設定
        onShape.Fill.ForeColor.RGB = RGB(cellColorCode Mod 256, Int(cellColorCode / 256) Mod 256, Int(cellColorCode / 256 / 256))
        
        '形の調整があればここに分岐をかく
        If flowList(i, TYPE_COL) = "ループ開始" Then
            onShape.AutoShapeType = 156
            onShape.Adjustments.Item(1) = 0.3
            onShape.Adjustments.Item(2) = 0
        ElseIf flowList(i, TYPE_COL) = "ループ終了" Then
            onShape.AutoShapeType = 156
            onShape.Adjustments.Item(1) = 0
            onShape.Adjustments.Item(2) = 0.3
        ElseIf flowList(i, TYPE_COL) = "参照" Then
            onShape.Height = 30
            onShape.Width = 30
        End If
    Next
End Function
Function moveFlowParts()
    '■2.場所移動
    Dim moveShape As Shape
    Dim top As Integer: top = Selection.top '100
    Dim nextShape As String
    Dim nextShapeArray As Variant
    '
    Dim nextNo As String
    '移動済みシェイプの名称をカンマ区切りで管理
    Dim movedShape As String
    
    '開始要素の移動
    Set moveShape = ActiveSheet.Shapes("1")
    '中央揃え
    moveShape.Left = center - moveShape.Width / 2
    moveShape.top = top
    movedShape = movedShape & nextShape & ","
    top = top + moveShape.Height + HEIGHT_MARGIN
    
    For i = LBound(flowList) To UBound(flowList)
        '遷移先情報の取得
        nextShape = CStr(flowList(i, AFTER_NO_COL))
        
        'ループ終了要素だったらEND遷移で単一処理のロジックを強引に通らせる
        If flowList(i, TYPE_COL) = "ループ終了" Then
            nextShape = Split(Split(nextShape, vbLf)(1), ":")(1)
        End If
        
        If nextShape Like "*" & vbLf & "*" Then
            '遷移先が複数ある場合
            nextShapeArray = Split(nextShape, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
                '遷移先のグループ
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                
                Set moveShape = ActiveSheet.Shapes(nextNo)
                    
                '中央揃え
                moveShape.Left = center - moveShape.Width / 2
                moveShape.top = top
                movedShape = movedShape & nextNo & ","
                '複数を横並びにさせるため
                moveShape.Left = center - (UBound(nextShapeArray) - j) * moveShape.Width * 1.3
                
                '案内テキストのシェイプ
                Set hjShape = ActiveSheet.Shapes.AddShape(61, 40, 10, 10, 10)
                hjShape.TextFrame.Characters.Text = Split(CStr(nextShapeArray(j)), ":")(0)
                hjShape.Name = moveShape.Name & "の補助"
                Set hjShape = hojoShape(hjShape)
                hjShape.Left = moveShape.Left - 30
                hjShape.top = moveShape.top - 20
            Next
            '次のtopプロパティ設定
            top = top + moveShape.Height + HEIGHT_MARGIN
        ElseIf Not movedShape Like "*" & nextShape & "*" Then
            '単一でかつ、未処理だったときだけ行う
            Set moveShape = ActiveSheet.Shapes(nextShape)
            '中央揃え
            moveShape.Left = center - moveShape.Width / 2
            moveShape.top = top
            movedShape = movedShape & nextShape & ","
            top = top + moveShape.Height + HEIGHT_MARGIN
        End If
    Next
    '終了要素の移動
    Set moveShape = ActiveSheet.Shapes(LAST_FLOW_NO)
    '中央揃え
    moveShape.Left = center - moveShape.Width / 2
    moveShape.top = top
End Function
Function addConnector()
    '■3.コネクタ付与
    Dim beforeShape As Shape
    Dim afterShape As Shape
    Dim beforeShapeName As String
    Dim afterShapeName As String
    Dim connectShape As Shape
    Dim startPoint As Integer
    Dim endPoint As Integer
    'エルボーコネクタが始点についてるシェイプ名称をカンマ区切りで管理
    Dim elbowNo As String
    For i = LBound(flowList) To UBound(flowList) - 1
        'beforeシェイプ
        beforeShapeName = CStr(flowList(i, FLOW_NO_COL))
        Set beforeShape = ActiveSheet.Shapes(beforeShapeName)
        startPoint = vlookup(shapeList, beforeShape.AutoShapeType, 4, 5)
        
        'afterシェイプ
        afterShapeName = CStr(flowList(i, AFTER_NO_COL))
        If afterShapeName Like "*" & vbLf & "*" Then
            '遷移先が複数ある場合
            nextShapeArray = Split(afterShapeName, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
                'コネクタ生成
                Set connectShape = setConnect(connectShape)
                
                'コネクタ接続（始点）
                connectShape.ConnectorFormat.BeginConnect beforeShape, startPoint
                
                '遷移先のグループ（ここ横並びのmoveもやっちゃうか）
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                
                'もしもbeforeシェイプが分岐系ならエルボーコネクタ
                If beforeShape.AutoShapeType = 63 Or beforeShape.AutoShapeType = 156 Then
                    connectShape.ConnectorFormat.Type = msoConnectorElbow
                    elbowNo = elbowNo & nextNo & ","
                End If
                
                '接続先シェイプの取得
                Set afterShape = ActiveSheet.Shapes(nextNo)
                '
                endPoint = vlookup(shapeList, afterShape.AutoShapeType, 4, 6)
                'コネクタ接続（終点）
                connectShape.ConnectorFormat.EndConnect afterShape, endPoint
                
                '「ループ終了」から「スープ開始」に伸びるコネクタの場合
                If Split(CStr(nextShapeArray(j)), ":")(0) = "Next" Then
                    connectShape.ConnectorFormat.BeginConnect beforeShape, 1
                    connectShape.ConnectorFormat.EndConnect afterShape, 1
                End If
                connectShape.Name = beforeShape.Name & "-" & afterShape.Name
            Next
        ElseIf i <> lastRow Then
            '接続先が単一で最終行でもない場合
            Set afterShape = ActiveSheet.Shapes(afterShapeName)
            
            Set connectShape = setConnect(connectShape)
            connectShape.ConnectorFormat.BeginConnect beforeShape, startPoint
            endPoint = vlookup(shapeList, afterShape.AutoShapeType, 4, 6)
            connectShape.ConnectorFormat.EndConnect afterShape, endPoint
            'こっちにいる
            If elbowNo Like "*" & beforeShapeName & "," & "*" Then
                connectShape.ConnectorFormat.Type = msoConnectorElbow
            End If
            connectShape.Name = beforeShape.Name & "-" & afterShape.Name
        End If
    
    Next
End Function
Function setCenterPosition(shp As Shape, center)
    shp.Left = center - shp.Width / 2
End Function
Function setConnect(connectShape As Shape)
    'コネクトシェイプ
    Set connectShape = ActiveSheet.Shapes.addConnector( _
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
Function vlookup(list As Variant, searchVal As Variant, searchCol As Integer, returnCol As Integer)
    vlookup = "61"
    For i = LBound(list) To UBound(list)
        If list(i, searchCol) Like "*" & searchVal & "*" Then
            vlookup = list(i, returnCol)
            Exit For
        End If
    Next
    
End Function

'コネクタの種類を直線とエルボーで切り替える
Sub B_コネクタ切替直線エルボー()
Attribute B_コネクタ切替直線エルボー.VB_ProcData.VB_Invoke_Func = "w\n14"

    If connectorErrorCheck Then
       Exit Sub
    End If
    
    With Selection.ShapeRange.ConnectorFormat
        If .Type = msoConnectorElbow Then
            .Type = msoConnectorStraight
        Else
            .Type = msoConnectorElbow
        End If
    End With
End Sub
'複数選択したシェイプのうち、最初のシェイプの左座標に他のシェイプの位置を合わせる
'TODO:これ本当はleft座標合わせじゃなくて1つ目のセンター合わせじゃないとあかんね
'だから・・・baseXCenterPointとbaseYCenterPointの2つが欲しい。まあ簡単に作れるか
Sub C_X座標合わせ()
Attribute C_X座標合わせ.VB_ProcData.VB_Invoke_Func = "e\n14"
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).Left + Selection.ShapeRange.Item(1).Width / 2
    For Each sp In Selection.ShapeRange
        sp.Left = baseXCenterPoint - sp.Width / 2
    Next
End Sub
Sub D_Y座標合わせ()
Attribute D_Y座標合わせ.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).top + Selection.ShapeRange.Item(1).Height / 2
    For Each sp In Selection.ShapeRange
        sp.top = baseXCenterPoint - sp.Height / 2
    Next
End Sub
'選択したコネクタの始点側のシェイプとの接続位置を変更する
Sub E_コネクタ始点変更()
Attribute E_コネクタ始点変更.VB_ProcData.VB_Invoke_Func = "r\n14"

    If connectorErrorCheck Then
       Exit Sub
    End If
    
    Dim currentBeginConnectPoint As Integer
    Dim targetShape As Shape
    
    With Selection.ShapeRange.ConnectorFormat
        '現在の始点接続ポイント取得
        currentBeginConnectPoint = .BeginConnectionSite
        '始点側のシェイプ取得
        Set targetShape = .BeginConnectedShape
        '次の接続ポイントに変更（現在がMAXの場合1に戻す）
        If currentBeginConnectPoint = targetShape.ConnectionSiteCount Then
            currentBeginConnectPoint = 1
        Else
            currentBeginConnectPoint = currentBeginConnectPoint + 1
        End If
        '接続ポイントを変更
        .BeginConnect targetShape, currentBeginConnectPoint
    End With
End Sub
'選択したコネクタの終点側のシェイプとの接続位置を変更する
Sub F_コネクタ終点変更()
Attribute F_コネクタ終点変更.VB_ProcData.VB_Invoke_Func = "t\n14"
    If connectorErrorCheck Then
       Exit Sub
    End If
    
    Dim currentEndConnectPoint As Integer
    Dim targetShape As Shape
    
    With Selection.ShapeRange.ConnectorFormat
        '現在の始点接続ポイント取得
        currentEndConnectPoint = .EndConnectionSite
        '始点側のシェイプ取得
        Set targetShape = .EndConnectedShape
        '次の接続ポイントに変更（現在がMAXの場合1に戻す）
        If currentEndConnectPoint = targetShape.ConnectionSiteCount Then
            currentEndConnectPoint = 1
        Else
            currentEndConnectPoint = currentEndConnectPoint + 1
        End If
        '接続ポイントを変更
        .EndConnect targetShape, currentEndConnectPoint
    End With
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
    If onShape.Width < 100 Then
        onShape.TextFrame2.AutoSize = msoAutoSizeNone
        onShape.Width = 100
    End If
    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    onShape.TextFrame2.WordWrap = msoTrue
    onShape.TextFrame2.AutoSize = msoAutoSizeNone
    onShape.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
    onShape.Height = 30
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
    If onShape.Width < 30 Then
        onShape.TextFrame2.AutoSize = msoAutoSizeNone
        onShape.Width = 30
    End If
    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    onShape.Height = 15
    Set hojoShape = onShape
End Function
'コネクタに処理をする際のエラーチェック。戻り値Trueでエラーあり、Falseでエラーなし
Function connectorErrorCheck()
    If TypeName(Selection) = "Range" Then
        MsgBox "コネクタを選んでください。"
        connectorErrorCheck = True
        Exit Function
    End If
    If Not Selection.ShapeRange.Connector Then
        MsgBox "コネクタを選んでください。"
        connectorErrorCheck = True
    End If
End Function
