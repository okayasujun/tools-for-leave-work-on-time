Attribute VB_Name = "main"
'処理No列
Const PROCESS_NO_COL = 1
'フローNo列
Const FLOW_NO_COL = 2
'遷移先列
Const DIST_NO_COL = 3
'種別列
Const TYPE_COL = 4
'シェイプテキスト列
Const SHAPE_TEXT_COL = 5
'[詳細]シートの走査開始行
Const LOOP_START_ROW = 2
'シェイプ縦間隔
Const HEIGHT_MARGIN = 32
'シェイプ横間隔
Const WIDTH_MARGIN = 130
'シェイプの幅
Const SHAPE_WIDTH = 130
'シェイプの高さ
Const SHAPE_HEIGHT = 33
'アニメーションフラグ
Const ANIMATION_FLAG = False
'最終フローNo（定数扱いなので大文字定義にする）
Dim LAST_FLOW_NO As String
'中央座標
Dim CENTER_POINT As Integer
'左列の中央座標
Dim LEFT_POINT As Integer
'シェイプリスト
Dim shapeList As Variant
'フローリスト
Dim flowList As Variant
'作成シェイプ
Dim onShape As shape
'Y座標
Dim yPoint As Integer
'移動済みシェイプ番号
Dim movedShapeNos As String
'標準シェイプ
'Dim baseShape As Shape
Sub A_フロー作成()
Attribute A_フロー作成.VB_ProcData.VB_Invoke_Func = "q\n14"

    '■1.初期設定
    Call init

    '■2.シェイプ作成
    Call createFlowParts

    '■3.場所移動
    Call moveFlowParts

    '■4.コネクタ付与
    Call addConnector

    '■5.調整
    Call adjust

End Sub
'■1.初期設定
Function init()

    '中央座標
     CENTER_POINT = Selection.Left + Selection.Width / 2
     LEFT_POINT = CENTER_POINT - 200

    'シェイプ種別値（AutoShapeType）を最適化
    With Sheets("シェイプ一覧")
    
        Dim lastShapeLine As Integer: lastShapeLine = .Cells(1, 1).End(xlDown).Row
        For i = 2 To lastShapeLine
            For Each shp In .Shapes
                If .Cells(i, 6).Left < shp.Left _
                    And .Cells(i, 6).Top < shp.Top _
                    And .Cells(i, 6).Top + .Cells(i, 6).Height > shp.Top + shp.Height Then

                    .Cells(i, 3) = shp.AutoShapeType

                    Exit For
                End If
            Next
        Next

        'シェイプリスト
        shapeList = .Range(.Cells(2, 1), .Cells(lastShapeLine, 5))
    End With
    
    With Sheets("詳細")

        Dim lastFlowLine As Integer: lastFlowLine = .Cells(1, 2).End(xlDown).Row

        'フローの最終番号
        LAST_FLOW_NO = .Cells(lastFlowLine, 2)

        'フローリスト（設計情報）
        flowList = .Range(.Cells(2, 1), .Cells(lastFlowLine, 5))
    
    End With
    yPoint = Selection.Top
End Function

'■2.シェイプ作成
Function createFlowParts()

    For i = LBound(flowList) To UBound(flowList)
        'シェイプ種別取得
        shapeType = vlookup(shapeList, flowList(i, TYPE_COL), 2, 3)
        'シェイプを生成
        Set onShape = ActiveSheet.Shapes.AddShape(shapeType, 40, 10, 100, 30)
        'テキストの設定
        onShape.TextFrame.Characters.text = flowList(i, PROCESS_NO_COL) & "." & flowList(i, SHAPE_TEXT_COL)
        '名前の設定（フローNo）
        onShape.Name = flowList(i, FLOW_NO_COL)
        'シェイプを最適化
        Set onShape = baseShape(onShape)

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
'■3.場所移動
Function moveFlowParts()
    '移動させるシェイプ
    Dim moveShape As shape
    Dim srcShapes As Variant
    Dim beforeCenterShape As String
    '開始要素の移動
    Call moveStartEnd(ActiveSheet.Shapes("1"))
    
    For i = LBound(flowList) + 1 To UBound(flowList) - 1
        Set moveShape = ActiveSheet.Shapes(CStr(flowList(i, FLOW_NO_COL)))
        '遷移元シェイプ名配列を取得
        srcShapes = getSrc(moveShape, CInt(i))
        '遷移元シェイプのうち、中央座標のものを取得
        beforeCenterShape = getMainPointShape(srcShapes)
        
        If moveShape.Name = "5" Then
            Debug.Print ""
        End If
        '※確率の低い分岐ほど前にもってくるのがいいんかな
        
        
        If UBound(srcShapes) = 0 And isSwitchShape(srcShapes, moveShape) <> "" Then 'ここの中でかすぎる。関数化してスマートに。
            '遷移元が1つのみかつ､それがSwitch要素
            Dim currentNo As Integer: currentNo = isSwitchShape(srcShapes, moveShape)
            If currentNo = 1 Then
                '走査シェイプが最初の要素のときのみトップを設定
                yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
            End If
            Dim branchCount As Integer: branchCount = getSwitchBranchCount(srcShapes) + 1
            '移動の重み
            Dim weight As Integer
            If branchCount Mod 2 = 0 Then
                '偶数
                Call switchEvenChildlenMove(moveShape, branchCount, currentNo)
            Else
                '奇数
                Call switchOddChildlenMove(moveShape, branchCount, currentNo)
            End If
            
'            moveShape.top = yPoint
            Call animationTop(moveShape, yPoint)
            
        ElseIf UBound(srcShapes) = 0 And isBranchShape(srcShapes, moveShape) Then
            '遷移元が１つのみかつ、かつ遷移元が分岐でその第一分岐先が検証中シェイプ（左移動する）
            'yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN 'どっちでもいいが、コネクタ処理にかかわる
'            moveShape.top = yPoint
            Call animationTop(moveShape, yPoint)
            Dim srtShapePoint As Integer: srtShapePoint = getCenterPoint(ActiveSheet.Shapes(srcShapes(0)))
'            moveShape.Left = srtShapePoint - moveShape.Width / 2 - 170 '遷移元シェイプとの差を指定する
            Call animationLeft(moveShape, srtShapePoint - moveShape.Width / 2 - 170)
        
        ElseIf UBound(srcShapes) = 0 And Not isCenterShape(srcShapes) Then
            '遷移元が１つのみかつ、それが中央ではない
            yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
'            moveShape.top = yPoint
'            moveShape.Left = CENTER_POINT - moveShape.Width / 2 - 170
            Call animationTop(moveShape, yPoint)
            Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 - 170)
            
        ElseIf UBound(srcShapes) = 0 And isCenterShape(srcShapes) Then
            '遷移元が１つのみかつ、それが中央
            yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
            
            Call animationTop(moveShape, yPoint)
            Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2)
            
        ElseIf UBound(srcShapes) > 0 Then  'And isCenterShape(srcShapes)
            '遷移元が2つ以上あり、その中にメイン座標のものがある（後ろの条件未実装）
            yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
'            moveShape.top = yPoint
'            moveShape.Left = CENTER_POINT - moveShape.Width / 2
            Call animationTop(moveShape, yPoint)
            Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2)
        Else
        End If
    Next
    
    '終了要素の移動
    Call moveStartEnd(ActiveSheet.Shapes(LAST_FLOW_NO))
End Function
Function switchEvenChildlenMove(moveShape As shape, branchCount As Integer, currentNo As Integer)
    Dim weight As Integer
    If branchCount / 2 >= currentNo Then
        '前半
        weight = branchCount / 2 - currentNo + 1
'       moveShape.Left = CENTER_POINT - moveShape.Width / 2 - (140 * weight) + 70
        Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 - (140 * weight) + 70)
    Else
        '後半
        weight = currentNo - branchCount / 2
'       moveShape.Left = CENTER_POINT - moveShape.Width / 2 + (140 * weight) - 70
        Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 + (140 * weight) - 70)
    End If
End Function
Function switchOddChildlenMove(moveShape As shape, branchCount As Integer, currentNo As Integer)
    If branchCount / 2 + 0.5 >= currentNo Then
        '前半+中央
        weight = branchCount / 2 + 0.5 - currentNo
'       moveShape.Left = CENTER_POINT - moveShape.Width / 2 - (140 * weight)
        Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 - (140 * weight))
    Else
        '後半
        weight = currentNo - (branchCount / 2 + 0.5)
'       moveShape.Left = CENTER_POINT - moveShape.Width / 2 + (140 * weight)
        Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 + (140 * weight))
    End If

End Function
Function animationTop(moveShape As shape, goalPoint As Integer)
    If ANIMATION_FLAG Then
        While Not isApproximate(moveShape.Top, goalPoint, 2)
            If isApproximate(moveShape.Top, goalPoint, 10) Then
                moveShape.Top = moveShape.Top + 1
            Else
                moveShape.Top = moveShape.Top + 9
            End If
            Application.wait [Now() + "0:00:00.0005"]
        Wend
    Else
        moveShape.Top = goalPoint
    End If
End Function
Function animationLeft(moveShape As shape, goalPoint As Integer)
    If ANIMATION_FLAG Then
        While moveShape.Left <> goalPoint
            If isApproximate(moveShape.Left, goalPoint, 10) Then
                moveShape.Left = moveShape.Left + 1
            Else
                moveShape.Left = moveShape.Left + 9
            End If
            Application.wait [Now() + "0:00:00.0005"]
        Wend
    Else
        moveShape.Left = goalPoint
    End If
End Function
Function isCenterShape(srcShapes As Variant)
    For Each no In srcShapes
        '近似値チェックにかける
        If isApproximate(getCenterPoint(ActiveSheet.Shapes(no)), CENTER_POINT) Then
            isCenterShape = True
            Exit Function
        End If
    Next
End Function
Function getCenterPoint(argShp As shape)
    getCenterPoint = argShp.Left + argShp.Width / 2
End Function
'Switch要素の分岐数を返す。シェイプリストのうち、最初に見つかったSwitch要素が対象になる
Function getSwitchBranchCount(srcShapes As Variant)
    For Each no In srcShapes
        distShape = vlookup(flowList, no, 2, 3)
        flowType = vlookup(flowList, no, 2, 4)
        distShapeArray = Split(distShape, vbLf)
        
        If flowType = "Switch" Then
            getSwitchBranchCount = UBound(distShapeArray)
        End If
    Next
End Function

Function isSwitchShape(srcShapes As Variant, moveShape As shape)
    Dim flowType As String
    Dim distShape As String
    Dim distShapeArray As Variant
    For Each no In srcShapes
        distShape = vlookup(flowList, no, 2, 3)
        flowType = vlookup(flowList, no, 2, 4)
        distShapeArray = Split(distShape, vbLf)
        
        If flowType = "Switch" Then
            For i = 0 To UBound(distShapeArray)
                If Split(distShapeArray(i), ":")(1) = moveShape.Name Then
                    isSwitchShape = i + 1
                    Exit Function
                End If
            Next
        End If
    Next
End Function
'指定されたシェイプ配列に分岐要素があるか、またその第一分岐先は今検証中のシェイプかどうか
Function isBranchShape(srcShapes As Variant, moveShape As shape)
    Dim flowType As String
    Dim distShape As String
    Dim distShapeArray As Variant
    For Each no In srcShapes
        '分岐
        distShape = vlookup(flowList, no, 2, 3)
        flowType = vlookup(flowList, no, 2, 4)
        distShapeArray = Split(distShape, vbLf)
        If flowType = "分岐" Then
            If Split(distShapeArray(0), ":")(1) = moveShape.Name Then
                isBranchShape = True
                Exit Function
            End If
        End If
    Next
End Function
'指定された遷移元シェイプのうち、メイン座標にあるもののシェイプ名を返す
Function getMainPointShape(srcShapes As Variant)
    For Each no In srcShapes
        '近似値チェック
        If isApproximate(CENTER_POINT, ActiveSheet.Shapes(no).Left + ActiveSheet.Shapes(no).Width / 2) Then
            mainStreetFlag = True
            getMainPointShape = no
            Exit Function
        End If
    Next
End Function
'遷移元のシェイプ名を
Function getSrc(moveShape As shape, currrentNo As Integer)
    Dim srcNos As String
    For j = LBound(flowList) To currrentNo
        If moveShape.Name = CStr(flowList(j, DIST_NO_COL)) Then
            '単一発見。遷移元格納
            srcNos = srcNos & CStr(flowList(j, FLOW_NO_COL)) & ","
    
        ElseIf CStr(flowList(j, DIST_NO_COL)) Like "*:" & moveShape.Name & vbLf & "*" _
            Or CStr(flowList(j, DIST_NO_COL)) Like "*:" & moveShape.Name Then
            '2桁の誤検知を防止するため、後ろの改行文字、前方のみ一致を条件とする
            '複数発見。遷移元格納
            srcNos = srcNos & CStr(flowList(j, FLOW_NO_COL)) & ","
        End If
    Next
    '配列で返す
    getSrc = Split(deleteEndText(srcNos), ",")
End Function
Function deleteEndText(text As String, Optional deleteLength As Long = 1) As String
    If Len(text) >= deleteLength Then
        deleteEndText = Left(text, Len(text) - deleteLength)
    Else
        deleteEndText = text
    End If
End Function
'開始、終了要素に対するシェイプ移動を施す
Function moveStartEnd(moveShape As shape)
    moveShape.Width = 75
'    moveShape.Left = CENTER_POINT - moveShape.Width / 2
            Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2)
    yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
'    moveShape.top = yPoint
            Call animationTop(moveShape, yPoint)
    movedShapeNos = movedShapeNos & moveShape.Name & ","
End Function
'
Function wait(waitTime As String)
    If ANIMATION_FLAG Then
        Application.wait [Now() + waitTime]
    End If
End Function
'■4.コネクタ付与
Function addConnector()
    Dim srcShape As shape
    Dim distShape As shape
    Dim connectShape As shape
    Dim startPoint As Integer
    Dim endPoint As Integer
    Dim hjShape As shape

    For i = LBound(flowList) To UBound(flowList) - 1
        Set srcShape = ActiveSheet.Shapes(CStr(flowList(i, FLOW_NO_COL)))
        startPoint = vlookup(shapeList, srcShape.AutoShapeType, 3, 4)
        
        If CStr(flowList(i, DIST_NO_COL)) Like "*" & vbLf & "*" Then
            '遷移先複数
            Dim distArray As Variant: distArray = Split(CStr(flowList(i, DIST_NO_COL)), vbLf)
            For j = LBound(distArray) To UBound(distArray)
                Set connectShape = baseConnect(connectShape)
                connectShape.ConnectorFormat.BeginConnect srcShape, startPoint
                
            If ANIMATION_FLAG Then
                Application.wait [Now() + "0:00:00.1"]
            End If

                Set distShape = ActiveSheet.Shapes(Split(CStr(distArray(j)), ":")(1))
                endPoint = vlookup(shapeList, distShape.AutoShapeType, 3, 5)
                connectShape.ConnectorFormat.EndConnect distShape, endPoint
                
                '中央座標をみてコネクタ種類を変更
                If Not isApproximate(srcShape.Left + srcShape.Width / 2, distShape.Left + distShape.Width / 2) Then
                    connectShape.ConnectorFormat.Type = msoConnectorElbow
                End If
                If isApproximate(srcShape.Top + srcShape.Height / 2, distShape.Top + distShape.Height / 2) Then
                    connectShape.ConnectorFormat.Type = msoConnectorStraight
                End If
                
                If UBound(distArray) = 1 And j = 0 Then
                    '分岐の1本目
                    connectShape.ConnectorFormat.BeginConnect srcShape, 2
                    '最終ポイントは右側にくるから接続ポイント数の指定とする
                    connectShape.ConnectorFormat.EndConnect distShape, distShape.ConnectionSiteCount
                End If
                connectShape.Name = srcShape.Name & "-" & distShape.Name
                '案内シェイプ
                Set hjShape = ActiveSheet.Shapes.AddShape(61, 40, 10, 10, 10)
                hjShape.TextFrame.Characters.text = Split(CStr(distArray(j)), ":")(0)
                hjShape.Name = connectShape.Name & "support"
                Set hjShape = supportShape(hjShape)
                If UBound(distArray) = 1 Then
                    '分岐
                    hjShape.Left = connectShape.Left + 2
                    hjShape.Top = connectShape.Top + 2
                Else
                    'Switch
                    hjShape.Left = distShape.Left + 63
                    hjShape.Top = distShape.Top - 16
                End If
            If ANIMATION_FLAG Then
                Application.wait [Now() + "0:00:00.1"]
            End If
            Next
        Else
            '遷移先単数
            Set distShape = ActiveSheet.Shapes(CStr(flowList(i, DIST_NO_COL)))
            Set connectShape = baseConnect(connectShape)
            
            connectShape.ConnectorFormat.BeginConnect srcShape, startPoint
            If ANIMATION_FLAG Then
                Application.wait [Now() + "0:00:00.1"]
            End If
            endPoint = vlookup(shapeList, distShape.AutoShapeType, 3, 5)
            connectShape.ConnectorFormat.EndConnect distShape, endPoint
            connectShape.Name = srcShape.Name & "-" & distShape.Name
            
            '中央座標をみてコネクタ種類を変更
            If Not isApproximate(srcShape.Left + srcShape.Width / 2, distShape.Left + distShape.Width / 2) Then
                connectShape.ConnectorFormat.Type = msoConnectorElbow
            End If
            
            '遷移先が「終了」フローだったら始点終点を変える (Not isApproximate(srcShape.Left + srcShape.Width / 2, distShape.Left + distShape.Width / 2)) And
            If distShape.Name = LAST_FLOW_NO Then
                If distShape.Top - srcShape.Top < 100 Then
                    connectShape.ConnectorFormat.BeginConnect srcShape, srcShape.ConnectionSiteCount - 1
                    connectShape.ConnectorFormat.EndConnect distShape, 1
                Else
                    '終了シェイプから遠い場合
                    connectShape.ConnectorFormat.BeginConnect srcShape, srcShape.ConnectionSiteCount - 2
                    connectShape.ConnectorFormat.EndConnect distShape, 2
                End If
            End If
            If ANIMATION_FLAG Then
                Application.wait [Now() + "0:00:00.1"]
            End If
        End If
    Next

End Function
'■5.調整
Function adjust()
    '全シェイプをみて線の重複がないかをチェックしたい
    'やり方がピンとこない
'    Dim beforeLeft As Integer
'    Dim leftest As Integer
'    Dim leftestShape As shape
'    For Each shp In ActiveSheet.Shapes
'        If shp.Connector Then
'
'            If shp.Left < beforeLeft Then
'                leftest = shp.Left
'                Set leftestShape = shp
'            End If
'            beforeLeft = shp.Left
'        End If
'    Next
End Function
'近似値かどうか
Function isApproximate(int1 As Integer, int2 As Integer, Optional diff As Integer = 1)
    If int1 < int2 Then
        isApproximate = int2 - int1 <= diff
    Else
        isApproximate = int1 - int2 <= diff
    End If
End Function
'シェイプ最適化
Function baseShape(onShape As shape)
    '■塗りつぶし（msoTrue:あり、msoFalse:なし）
    onShape.Fill.Visible = msoTrue
    '■線の太さ。お好みでどうぞ
    onShape.Line.weight = 1
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
    'onShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    'セルに合わせて移動やサイズ変更をしない
    onShape.Placement = xlFreeFloating
    '高さ
    onShape.Height = SHAPE_HEIGHT
    '幅
    onShape.Width = SHAPE_WIDTH

    If onShape.Width < SHAPE_WIDTH Then
        onShape.TextFrame2.AutoSize = msoAutoSizeNone
    Else
        'onShape.Width = SHAPE_HEIGHT + 20
        'onShape.Width = onShape.Width + 20
    End If

    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    onShape.TextFrame2.WordWrap = msoTrue
    onShape.TextFrame2.AutoSize = msoAutoSizeNone
    onShape.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
    'TODO:セルが縦に広がっても、変わらないオプションをつけること
    Set baseShape = onShape
End Function
'案内シェイプ
Function supportShape(onShape As shape)
    '■塗りつぶし（msoTrue:あり、msoFalse:なし）
    onShape.Fill.Visible = msoFalse
    '■線の太さ。お好みでどうぞ
    onShape.Line.weight = 0
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
    Set supportShape = onShape
End Function
'コネクタ
Function baseConnect(connectShape As shape)
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
    connectShape.Line.weight = 1
    Set baseConnect = connectShape
End Function
'表から使用したいシェイプを探す
Function vlookup(list As Variant, searchVal As Variant, searchCol As Integer, returnCol As Integer)
    '初期値
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
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).Top + Selection.ShapeRange.Item(1).Height / 2
    For Each sp In Selection.ShapeRange
        sp.Top = baseXCenterPoint - sp.Height / 2
    Next
End Sub
'選択したコネクタの始点側のシェイプとの接続位置を変更する
Sub E_コネクタ始点変更()
Attribute E_コネクタ始点変更.VB_ProcData.VB_Invoke_Func = "r\n14"

    If connectorErrorCheck Then
       Exit Sub
    End If

    Dim currentBeginConnectPoint As Integer
    Dim targetShape As shape

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
    Dim targetShape As shape

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
