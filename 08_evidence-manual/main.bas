Attribute VB_Name = "main"
'###############################
'機能名：エビデンス・マニュアル作成支援ツール v2.0
'Author：okayasu jun
'作成日：2022/10/19
'更新日：2023/02/25
'COMMENT：
'###############################
'ポインタAPI。マウスカーソル位置からセル位置を取得するために使用する
Private Type POINTAPI
    x As Long
    y As Long
End Type
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'■画像タイトルに補記をする場合の行数（タイトルと画像の間の空行数）
Const REMARK_LINE = 0
'#選択セル範囲の大きさの赤枠をマウス位置に出現させる
Sub A_赤枠を出現させる()
Attribute A_赤枠を出現させる.VB_ProcData.VB_Invoke_Func = "q\n14"

    Dim onShape As shape
    Dim beforeRange As Range
    
 On Error GoTo ErrHndl
    '処理前後で選択セル位置を保持するため
    Set beforeRange = Selection
    
    'シェイプ生成&スタイル設定
    '■第一引数は以下のURLを参考に変更可。シェイプの形を指定する
    'https://learn.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype?redirectedfrom=MSDN
    Set onShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, _
                                            Selection.left, _
                                            Selection.top, _
                                            Selection.width, _
                                            Selection.height)
    '■塗りつぶし（msoTrue:あり、msoFalse:なし）
    onShape.Fill.Visible = msoFalse
    '■線の太さ。お好みでどうぞ
    onShape.Line.Weight = 4
    '■色指定
    onShape.Line.ForeColor.RGB = RGB(255, 0, 0)
    '■塗りつぶし色指定
    'onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    '画像の上にカーソルがある状態では後続処理ができないため、一時非表示にする
    'この段階でどのシェイプ上にカーソルがあるか不明なので全シェイプを対象とする
    For Each shp In ActiveSheet.Shapes
        shp.Visible = False
    Next
    
    '座標取得処理のために、マウスカーソルの状態を解放する
    beforeRange.Select
    
    '以下で上記作成シェイプをマウスがあるセルの位置に移動させる
    Dim p As POINTAPI
    Dim Getcell As Range

    'カーソル位置を取得
    GetCursorPos p

    'マウスカーソルの位置からセルを取得（カーソルの状態次第では失敗する）
    Set Getcell = ActiveWindow.RangeFromPoint(p.x, p.y)

    'シェイプ位置をマウスカーソルの直近セルの左上に合わせる
    onShape.top = Getcell.top
    onShape.left = Getcell.left
    
    '全シェイプを可視状態に戻す
    For Each shp In ActiveSheet.Shapes
        shp.Visible = True
    Next
    Exit Sub
ErrHndl:
    'MsgBox "エラー発生"
    '全シェイプを可視状態に戻す
    For Each shp In ActiveSheet.Shapes
        shp.Visible = True
    Next

End Sub
'#コネクタ系以外のシェイプ及び画像に影をつける（塗りつぶしなしのオートシェイプは対象外）
'グループはその構成シェイプすべてに影がついてしまうから条件にしない
Sub B_影を付ける()
Attribute B_影を付ける.VB_ProcData.VB_Invoke_Func = "w\n14"

    Dim shp As shape

    If TypeName(Selection) = "Range" Then
        'シェイプ未選択状態。全シェイプを対象にする
        For Each shp In ActiveSheet.Shapes
            '条件変更時参考：https://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
            If shp.Type = msoPicture Or (shp.Type = msoAutoShape And shp.Fill.Visible) Then
                Call castShadow(shp)
            End If
        Next
    Else
        '選択中シェイプのみ処理を行う
        For Each shp In Selection.ShapeRange
            If shp.Type = msoPicture Or (shp.Type = msoAutoShape And shp.Fill.Visible) Then
                Call castShadow(shp)
            End If
        Next
    End If
End Sub
'指定されたシェイプに影を付与します。
Function castShadow(shp As shape)
    With shp.Shadow
        '■影の種類：https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.core.msoshadowtype?view=office-pia
        .Type = msoShadow26
        '影の表示切替
        .Visible = msoTrue
        '■影の効果：https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.core.msoshadowstyle?view=office-pia
        .Style = msoShadowStyleOuterShadow
        '■ブロアー。影のぼかし具合
        .Blur = 20
        '■影の相対位置
        .OffsetX = 7.7781745931
        .OffsetY = 7.7781745931
        '■影をシェイプとともに回転させるかどうか
        .RotateWithShape = msoFalse
        '■影の色
        .ForeColor.RGB = RGB(100, 100, 100)
        '■トランスパレンシー。影の透明度。0～1で指定。1は完全に透明
        .Transparency = 0.4
        '影のサイズ
        .Size = 100
    End With
End Function
'#画像の効果をリセットする
Sub C_効果をリセットする()
Attribute C_効果をリセットする.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim shp As shape
    If TypeName(Selection) = "Range" Then
        'シェイプ未選択状態。全シェイプを対象にする
        For Each shp In ActiveSheet.Shapes
            Call shapeReset(shp)
        Next
    Else
        '選択中シェイプのみ処理を行う
        For Each shp In Selection.ShapeRange
            Call shapeReset(shp)
        Next
    End If
End Sub
'指定されたシェイプ及び画像の効果を初期化します。
'コネクタ系以外のシェイプ及び画像をリセットする（塗りつぶしなしのオートシェイプは対象外）
Function shapeReset(shp As shape)
    If shp.Type = msoPicture Or (shp.Type = msoAutoShape And shp.Fill.Visible) Then
        With Application.CommandBars
            '「図のリセット」が可能なときのみ
            If .GetEnabledMso("PictureReset") Then
                .ExecuteMso "PictureReset"
            End If
            'オートシェイプの場合は以下で外す（■ケースに応じて初期化処理を追加すること）
            shp.Shadow.Visible = msoFalse
        End With
    End If
End Function
'#シェイプを整列させる（貼付順）
Sub D_貼付順に整列させる()
Attribute D_貼付順に整列させる.VB_ProcData.VB_Invoke_Func = "e\n14"
    '■画像間の間隔
    Const MARGIN_BOTTOM = 70
    
    
    '貼付座標を格納する（topは都度書き換え、leftは初期値を使いまわす）
    Dim top As Integer: top = Selection.top + 5
    
    'キャプションを記載する用のセル
    Dim captionRange As Range
    Dim moveShape As shape
    
    'エラーチェック
    If Selection.Row - REMARK_LINE - 1 < 1 Then
        MsgBox "キャプション用の行が足りません。あと" & REMARK_LINE - Selection.Row + 2 & "行下の位置で実行してください。"
        Exit Sub
    End If
    
    'キャプションタイトル
    Dim captionText As String
    '■ダイアログを使う場合は以下のコメントアウト分を使用する
    captionText = "▼" 'InputBox("キャプションの初期値を入れて。", "キャプションオプション", "▼ここに画像の説明を書く")
    
    If StrPtr(answer) = 0 Then
        'キャンセル時
        Exit Sub
    End If
    
    For Each moveShape In ActiveSheet.Shapes
        '次に該当しないものは対象外：画像、グループ、塗りつぶしのないオートシェイプ
        '条件変更時参考：https://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
        If moveShape.Type <> msoPicture _
            And moveShape.Type <> msoGroup _
            And (moveShape.Type = msoAutoShape And Not moveShape.Fill.Visible) Then
            GoTo CONTINUE:
        End If
        
        'シェイプを移動させて、
        Set captionRange = move(moveShape, top)
        
        '■キャプション入力の設定（不要ならコメントアウトして）
        Call setCaption(captionRange, captionText)
        
        '今対象にしたシェイプの上部座標 + 今対象にしたシェイプの高さ + 画像間の間隔 + キャプションセル行の高さ = 次のシェイプの移動先上部座標
        top = top + moveShape.height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).height
CONTINUE:
    Next
    
    'END処理
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
    Call setCaption(dummyShape.TopLeftCell, "END")
    dummyShape.Delete
    
End Sub
'
Function move(moveShape As shape, top As Integer)
    '移動位置を取得するためのダミーシェイプ
    Dim dummyShape As shape
    Dim left As Integer: left = Selection.left
    
    '左上隅のセルを取得するためのダミーシェイプ
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
        
    'シェイプを移動する
    moveShape.top = dummyShape.TopLeftCell.Offset(0, 0).top
    moveShape.left = Selection.left
        
    'キャプション入力用セルを取得する（-1はタイトル分）
    Set move = dummyShape.TopLeftCell.Offset(-1 - REMARK_LINE, 0)
        
    '用済みだから削除する
    dummyShape.Delete
End Function
'キャプション用セルの設定
Function setCaption(captionRange As Range, captionText As String)
    '画像間移動をCtrl+矢印で高速に行うため
    captionText = IIf(captionText = "", " ", captionText)
    '■適宜変えてよし
    captionRange.Value = captionText
    captionRange.Font.Bold = True
    captionRange.Font.Color = RGB(0, 0, 0)
End Function
'選択中のシェイプを選択順にコネクタで繋ぐ
Sub F_シェイプを選択順にコネクタで繋ぐ()
Attribute F_シェイプを選択順にコネクタで繋ぐ.VB_ProcData.VB_Invoke_Func = "l\n14"
    Dim startShape As shape
    Dim endShape As shape
    Dim connectShape As shape
    
    If TypeName(Selection) = "Range" Then
        MsgBox "シェイプが選択されていません。2つ以上選択してください。"
        Exit Sub
    End If
    For Each shp In Selection.ShapeRange
        If shp.Type = msoGroup Or shp.Connector Then
            MsgBox "選択シェイプにグループかコネクタが含まれています。解除してください。"
            Exit Sub
        End If
    Next
    
    For i = 1 To Selection.ShapeRange.Count - 1
        '選択中シェイプの保持（接続元）
        Set startShape = Selection.ShapeRange.Item(i)
        '選択中シェイプの保持（接続先）
        Set endShape = Selection.ShapeRange.Item(i + 1)

        '接続シェイプの誕生
        '■Type引数は右記を参照：https://learn.microsoft.com/ja-jp/office/vba/api/office.msoconnectortype
        Set connectShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorElbow, BeginX:=0, BeginY:=0, EndX:=0, EndY:=0)
        '■接続の始点位置指定（最後の引数は1:上辺、2:左辺、3:下辺、4右辺）
        connectShape.ConnectorFormat.BeginConnect ActiveSheet.Shapes(startShape.Name), 3
        '■接続の終点位置指定（最後の引数は始点位置の指定と同様）
        connectShape.ConnectorFormat.EndConnect ActiveSheet.Shapes(endShape.Name), 2
        '■終点コネクタを三角に。
        connectShape.Line.EndArrowheadStyle = msoArrowheadTriangle
        '■線の色
        connectShape.Line.ForeColor.RGB = RGB(0, 0, 0)
        '■線の太さ
        connectShape.Line.Weight = 1
        '■終点の長さ
        connectShape.Line.EndArrowheadLength = msoArrowheadLong
        '■終点の太さ
        connectShape.Line.EndArrowheadWidth = msoArrowheadWide
    Next
End Sub
'コネクタの種類を直線とエルボーで切り替える
Sub G_コネクタ種類切り替え()
Attribute G_コネクタ種類切り替え.VB_ProcData.VB_Invoke_Func = "i\n14"

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
'選択したコネクタの始点側のシェイプとの接続位置を変更する
Sub H_コネクタ始点変更()
Attribute H_コネクタ始点変更.VB_ProcData.VB_Invoke_Func = "o\n14"

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
Sub I_コネクタ終点変更()
Attribute I_コネクタ終点変更.VB_ProcData.VB_Invoke_Func = "p\n14"
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
'#選択中のシェイプをグループ化
Sub J_選択中シェイプをグループ化()
    Selection.Group.Select
End Sub
'#選択中のシェイプをグループ解除
Sub K_選択中シェイプをグループ解除()
    Selection.Ungroup
End Sub
'選択中のシェイプを最背面にする
Sub L_選択中シェイプを最背面へ()
    If TypeName(Selection) = "Range" Then
        MsgBox "シェイプを一つ選んでから実行してね。"
        Exit Sub
    End If
    
    For Each shp In Selection.ShapeRange
        shp.ZOrder msoSendToBack
    Next
End Sub
'#シェイプの場所に値がなくなるように空行を挿入する
Sub N_シェイプ貼付時､ブランク行挿入()
    '#クリップボードにデータがある時のみ
    If Application.ClipboardFormats(1) Then
        '貼付。Ctrl + Vにあたるアクション（この時点でSelectionはシェイプになる模様）
        ActiveSheet.Paste
        '移動位置を取得するためのダミーシェイプ
        Dim dummyShape As shape
        '左下隅のセルを取得するためのダミーシェイプ
        Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, Selection.top + Selection.height, 1, 1)
    
        '「セルに合わせて移動やサイズ変更をしない」に設定
        'これやらないと行の挿入に合わせてシェイプも一緒に伸びてしまうから
        Selection.Placement = xlFreeFloating
        
        '貼付シェイプの下にあるセル分ループ
        For i = Selection.TopLeftCell.Row To dummyShape.TopLeftCell.Row
            '列ループ（■j=Selection.TopLeftCell.Columnならシェイプ貼付位置から開始）
            For j = 1 To 100
                '対象行のどこかに値があれば行を挿入する
                If Cells(i, j) <> "" Then
                    Rows(i).Insert
                    Exit For
                End If
            Next
        Next
    
        '用済みだから削除する
        dummyShape.Delete
    End If
End Sub
'#先頭に目次のシートを作成する
Sub O_目次シートを作成する()
    Dim ws As Worksheet
    
    '関数は別途参照
    If Not isExistCheckToSheet(ThisWorkbook, "目次") Then
        Worksheets.Add before:=Sheets(1)
        Set ws = Sheets(1)
        ws.Name = "目次"
        ws.Cells(1, 1) = "No."
        ws.Cells(1, 2) = "シート名"
        ws.Cells(1, 3) = "シートの説明"
        ws.Cells(1, 4) = "シェイプの数"
        ws.Cells(1, 5) = "使用範囲"
        ws.Cells(1, 6) = "備考"
        ws.Cells(1, 7) = "作成者"
        ws.Cells(1, 8) = "作成日"
        'フォント色
        Range("A1:H1").Font.Color = RGB(20, 10, 10)
        '背景色
        Range("A1:H1").Interior.Color = RGB(255, 242, 204)
        '太字
        Range("A1:H1").Font.Bold = True
        Cells(2, 1).Select
        'ウィンドウ枠の固定
        ActiveWindow.FreezePanes = True
        '目盛線非表示
        ActiveWindow.DisplayGridlines = False
    Else
        Set ws = Sheets(1)
    End If
    
    Dim loopWs As Worksheet
    
    For i = 2 To Worksheets.Count
        Set loopWs = Worksheets(i)
        ws.Cells(i, 1) = i - 1
        ws.Cells(i, 2) = loopWs.Name
        ws.Cells(i, 4) = loopWs.Shapes.Count
        ws.Cells(i, 5) = loopWs.UsedRange.Address
    Next
    ws.Columns("A:H").AutoFit
End Sub
'指定されたブックに指定されたシートが存在するかどうかを返す。存在する:TRUE、存在しない:FALSE
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
        isExistCheckToSheet = checkSheet <= wb.Worksheets.Count
    End If
End Function
'#アクティブシートの内容に従いシートを生成し、リンクを付与する。ソートも行う
Sub P_シート生成とリンク付与()
    'TODO:目盛線削除、縮尺統一
    Dim topSheet As Worksheet
    Set topSheet = ActiveSheet
    
    Dim lastRowToBottom As Integer: lastRowToBottom = topSheet.Cells(1, 2).End(xlDown).Row
    
    Dim sheetName As String
    Dim linkRange As Range

    For i = 2 To lastRowToBottom
        sheetName = topSheet.Cells(i, 2).Value
        Set linkRange = topSheet.Cells(i, 2)
        
        If Not existsSheet(sheetName) Then
            'シートが存在していない場合
            With Worksheets.Add(after:=ActiveSheet)
                .Name = sheetName
                topSheet.Hyperlinks.Add Anchor:=linkRange, Address:="", SubAddress:=.Name & "!A1"
                .Select
            End With
        Else
            '既にある場合
            topSheet.Hyperlinks.Add Anchor:=linkRange, Address:="", SubAddress:=sheetName & "!A1"
            Sheets(sheetName).Select
            'シート順の並び変え
            If existsSheet(topSheet.Cells(i - 1, 2)) Then
                'ここ、シート名の指定に「.value」が必要
                Sheets(sheetName).move after:=Sheets(topSheet.Cells(i - 1, 2).Value)
            End If
        End If
    Next
    
    topSheet.Select
End Sub
'シートが存在するかどうか
Function existsSheet(ByVal sheetName As String)
    Dim ws As Variant
    For Each ws In Sheets
        If LCase(ws.Name) = LCase(sheetName) Then
            existsSheet = True
            Exit Function
        End If
    Next

    '存在しない
    existsSheet = False
End Function
'シェイプのうち選択中セルのleftプロパティに一致しないものを選択セル位置から並べる
Sub Q_シェイプ追加整列()
Attribute Q_シェイプ追加整列.VB_ProcData.VB_Invoke_Func = " \n14"
    '■画像間の間隔
    Const MARGIN_BOTTOM = 70
    
    
    '貼付座標を格納する（topは都度書き換え、leftは初期値を使いまわす）
    Dim top As Integer: top = Selection.top + 5
    
    'キャプションを記載する用のセル
    Dim captionRange As Range
    Dim moveShape As shape 'キャプションタイトル
    Dim captionText As String: captionText = "▼"
    For Each moveShape In ActiveSheet.Shapes
        '次に該当しないものは対象外：画像、グループ、塗りつぶしのないオートシェイプ、
        'もしくは選択中セルleftプロパティと対象シェイプleftプロパティが一致しないもの
        If (moveShape.Type <> msoPicture _
            And moveShape.Type <> msoGroup _
            And (moveShape.Type = msoAutoShape And Not moveShape.Fill.Visible)) _
            Or moveShape.left = Selection.left Then
            GoTo CONTINUE:
        End If
        
        'シェイプを移動させて、
        Set captionRange = move(moveShape, top)
        
        '■キャプション入力の設定（不要ならコメントアウトして）
        Call setCaption(captionRange, captionText)
        
        '今対象にしたシェイプの上部座標 + 今対象にしたシェイプの高さ + 画像間の間隔 + キャプションセル行の高さ = 次のシェイプの移動先上部座標
        top = top + moveShape.height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).height
CONTINUE:
    Next
End Sub
Sub R_2か3カラムの並び()
    'また今度実装する
End Sub
Sub U_セルからセルに伸びるコネクタ()
    'また今度実装する
End Sub
'たたき台的クオリティ。過不足は後日修正
Sub V_最初のシェイプをスキャンコピー()
    '複数シェイプ選択時、2つ目以降は1つ目のスタイルを適用する。ループ処理
    Dim baseShp As shape
    Dim shp As shape
    Set baseShp = Selection.ShapeRange.Item(1)
    For i = 2 To Selection.ShapeRange.Count
        '選択中シェイプの保持（接続元）
        Set shp = Selection.ShapeRange.Item(i)
        shp.Line.ForeColor.RGB = baseShp.Line.ForeColor.RGB
        'shp.ForeColor.RGB = baseShp.ForeColor.RGB
        '↓ワードアートフォーマットが指定できないシェイプを選ぶときはコメントアウトして
        shp.TextFrame2.WordArtformat = baseShp.TextFrame2.WordArtformat
        shp.Fill.Transparency = baseShp.Fill.Transparency
        'テキストまで変えたくないときはコメントアウト
        'shp.TextFrame.Characters.Text = baseShp.TextFrame.Characters.Text
        shp.Fill.ForeColor.RGB = baseShp.Fill.ForeColor.RGB
        shp.TextFrame2.TextRange.Font.Size = baseShp.TextFrame2.TextRange.Font.Size
        shp.TextFrame2.WordWrap = baseShp.TextFrame2.WordWrap
        shp.TextFrame.Characters.Font.Color = baseShp.TextFrame.Characters.Font.Color
        shp.TextFrame.Characters.Font.Name = baseShp.TextFrame.Characters.Font.Name
        shp.TextFrame2.VerticalAnchor = baseShp.TextFrame2.VerticalAnchor
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = baseShp.TextFrame2.TextRange.ParagraphFormat.Alignment
        shp.Placement = baseShp.Placement
        shp.LockAspectRatio = baseShp.LockAspectRatio
        shp.TextFrame2.AutoSize = baseShp.TextFrame2.AutoSize
        shp.TextFrame2.MarginLeft = baseShp.TextFrame2.MarginLeft
        shp.TextFrame2.MarginRight = baseShp.TextFrame2.MarginRight
        shp.TextFrame2.MarginTop = baseShp.TextFrame2.MarginTop
        shp.TextFrame2.MarginBottom = baseShp.TextFrame2.MarginBottom
        shp.TextFrame2.WordWrap = baseShp.TextFrame2.WordWrap
        shp.TextFrame2.VerticalAnchor = baseShp.TextFrame2.VerticalAnchor
        shp.TextFrame2.HorizontalAnchor = baseShp.TextFrame2.HorizontalAnchor
        shp.TextFrame2.Orientation = baseShp.TextFrame2.Orientation
    Next
End Sub
