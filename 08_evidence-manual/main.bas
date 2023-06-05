Attribute VB_Name = "main"
'###############################
'機能名：エビデンス・マニュアル作成支援ツール v2.0
'Author：okayasu jun
'作成日：2022/10/19
'更新日：2023/05/13
'COMMENT：各コメントの「■」は変更可能を示す。用途や好みに合わせて変えてみて。
'###############################
'ポインタAPI。マウスカーソル位置からセル位置を取得するために使用する
Private Type POINTAPI
    x As Long
    y As Long
End Type
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'■画像キャプションに補記をする場合の行数（画像とキャプションの間の空行数）
Const REMARK_LINE = 0
Const CAPTION_TEXT_TOP_FLAG = True
'#選択セル範囲の大きさの赤枠をマウス位置に出現させる
Sub AA_赤枠を出現させる()
Attribute AA_赤枠を出現させる.VB_ProcData.VB_Invoke_Func = "q\n14"

    '赤枠シェイプを格納する変数
    Dim onShape As Shape
    '処理開始時に選択しているセル情報
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
                                            Selection.WIDTH, _
                                            Selection.Height)
    '■塗りつぶし（msoTrue:あり、msoFalse:なし）
    onShape.Fill.Visible = msoFalse
    '■線の太さ
    onShape.Line.Weight = 4
    '■色指定
    onShape.Line.ForeColor.RGB = RGB(255, 0, 0)
    '■塗りつぶし色指定
    'onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    '■線のスタイル（実線/点線）
    'onShape.Line.DashStyle = msoLineDash
    '■線のスタイル（一重線/二重線）
    'onShape.Line.Style = msoLineThinThin
    
    '■選択セル位置に表示するだけでいい場合はここで処理を終了する。コメントインして
    'Exit Sub
    
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

    'マウスカーソルの位置からセルを取得
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
Sub AB_影を付ける()
Attribute AB_影を付ける.VB_ProcData.VB_Invoke_Func = "w\n14"

    Dim shp As Shape

    If TypeName(Selection) = "Range" Then
        'シェイプ未選択状態。全シェイプを対象にする
        For Each shp In ActiveSheet.Shapes
            '条件変更したいときはこちらを参考に：https://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
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
'指定されたシェイプに影を付する
Function castShadow(shp As Shape)
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
Sub AC_効果をリセットする()
Attribute AC_効果をリセットする.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim shp As Shape
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
Function shapeReset(shp As Shape)
    If shp.Type = msoPicture Or (shp.Type = msoAutoShape And shp.Fill.Visible) Then
        With Application.CommandBars
            '「図のリセット」が可能なときのみ
            If .GetEnabledMso("PictureReset") Then
                .ExecuteMso "PictureReset"
            End If
            'オートシェイプの場合は以下で外す（■やりたいことに応じて処理内容を追加すること）
            shp.Shadow.Visible = msoFalse
        End With
    End If
End Function
'#シェイプを整列させる（貼付順）
Sub AD_シェイプを貼付順に整列させる()
Attribute AD_シェイプを貼付順に整列させる.VB_ProcData.VB_Invoke_Func = "e\n14"
    '■画像間の間隔
    Const MARGIN_BOTTOM = 70
    
    '貼付座標を格納する（topは都度書き換え、leftは初期値を使いまわす）
    Dim top As Integer: top = Selection.top + 5
    
    'キャプションを記載する用のセル
    Dim captionRange As Range
    Dim moveShape As Shape
    
    'エラーチェック
    If Selection.Row - REMARK_LINE - 1 < 1 Then
        MsgBox "キャプション用の行が足りません。あと" & REMARK_LINE - Selection.Row + 2 & "行下の位置で実行してください。"
        Exit Sub
    End If
    
    'キャプションタイトル
    Dim captionText As String
    '■ダイアログを使う場合は以下のコメントアウト部分を使用する
    If CAPTION_TEXT_TOP_FLAG Then
        captionText = "▼" 'InputBox("キャプションの初期値を入れて。", "キャプションオプション", "▼ここに画像の説明を書く")
    Else
        captionText = "▲" 'InputBox("キャプションの初期値を入れて。", "キャプションオプション", "▲ここに画像の説明を書く")
    End If
    
    If StrPtr(captionText) = 0 Then
        'キャンセル時
        Exit Sub
    End If
    
    For Each moveShape In ActiveSheet.Shapes
        '次に該当するものが対象：画像、グループ（■必要に応じて条件調整して）
        '条件変更時参考：https://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
        If moveShape.Type <> msoPicture _
            And moveShape.Type <> msoGroup _
            And (moveShape.Type = msoAutoShape) Then 'And Not moveShape.Fill.Visible
            GoTo CONTINUE:
        End If
        
        'シェイプを移動させて、
        Set captionRange = move(moveShape, top)
        
        '■キャプション入力の設定（不要ならコメントアウトして）
        Call setCaption(captionRange, captionText)
        
        '今対象にしたシェイプの上部座標 + 今対象にしたシェイプの高さ + 画像間の間隔 + キャプションセル行の高さ = 次のシェイプの移動先上部座標
        top = top + moveShape.Height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).Height
CONTINUE:
    Next
    
    'END処理
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
    Call setCaption(dummyShape.TopLeftCell, "- END -")
    dummyShape.Delete
    
End Sub
'指定されたシェイプを第二引数の位置に移動させる
Function move(moveShape As Shape, top As Integer)
    '移動位置を取得するためのダミーシェイプ
    Dim dummyShape As Shape
    Dim left As Integer: left = Selection.left
    
    '左上隅のセルを取得するためのダミーシェイプ
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
        
    'シェイプを移動する
    moveShape.top = dummyShape.TopLeftCell.Offset(0, 0).top
    moveShape.left = Selection.left
    
    '作業記録の意味合いで画像名に時間を設定する。一意にするためミリ秒も末尾に付与
    moveShape.Name = "image-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
    
    If CAPTION_TEXT_TOP_FLAG Then
        'キャプション入力用セルを取得する（-1はタイトル分）
        Set move = dummyShape.TopLeftCell.Offset(-1 - REMARK_LINE, 0)
    Else
        dummyShape.top = moveShape.top + moveShape.Height
        'キャプション入力用セルを取得する（-1はタイトル分）
        Set move = dummyShape.TopLeftCell.Offset(1, 0)
    End If
    
    '用済みだから削除する
    dummyShape.Delete
End Function
'キャプション用セルの設定
Function setCaption(captionRange As Range, captionText As String)
    '画像間移動をCtrl+矢印で高速に行うため
    captionText = IIf(captionText = "", " ", captionText)
    '■適宜変えてよし。お好みで
    captionRange.Value = captionText
    captionRange.Font.Bold = True
    captionRange.Font.Color = RGB(40, 40, 40)
    captionRange.Font.Size = 18
    captionRange.Font.Name = "BIZ UDPゴシック" '"Meiryo UI"'
End Function
'ミリ秒を取得
Function getMSec() As String
    Dim dblTimer As Double
    Dim s_return As String
    dblTimer = CDbl(Timer)
    s_return = Format(Fix((dblTimer - Fix(dblTimer)) * 1000), "000")
    getMSec = s_return
End Function
'選択中のシェイプを選択順にコネクタで繋ぐ
Sub AF_シェイプを選択順にコネクタで繋ぐ()
Attribute AF_シェイプを選択順にコネクタで繋ぐ.VB_ProcData.VB_Invoke_Func = "l\n14"
    Dim startShape As Shape
    Dim endShape As Shape
    Dim connectShape As Shape
    
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
    
    For i = 1 To Selection.ShapeRange.count - 1
        '選択中シェイプの保持（接続元）
        Set startShape = Selection.ShapeRange.Item(i)
        '選択中シェイプの保持（接続先）
        Set endShape = Selection.ShapeRange.Item(i + 1)

        '接続シェイプの誕生
        '■Type引数は右記を参照：https://learn.microsoft.com/ja-jp/office/vba/api/office.msoconnectortype
        Set connectShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorStraight, BeginX:=0, BeginY:=0, EndX:=0, EndY:=0)
        '■接続の始点位置指定（最後の引数は1:上辺、2:左辺、3:下辺、4右辺）
        connectShape.ConnectorFormat.BeginConnect ActiveSheet.Shapes(startShape.Name), 4
        '■接続の終点位置指定（最後の引数は始点位置の指定と同様）
        connectShape.ConnectorFormat.EndConnect ActiveSheet.Shapes(endShape.Name), 2
        'コネクタを加工する
        Call makeConnectAllow(connectShape)
    Next
End Sub
'コネクタを加工する。各プロパティ好みに合わせて設定されたし
Function makeConnectAllow(connectShape As Shape)
    '■終点コネクタを三角に。
    connectShape.Line.EndArrowheadStyle = msoArrowheadTriangle
    '■線の色
    connectShape.Line.ForeColor.RGB = RGB(10, 10, 10)
    '■線の太さ
    connectShape.Line.Weight = 1
    '■終点の長さ
    connectShape.Line.EndArrowheadLength = msoArrowheadLong
    '■終点の太さ
    connectShape.Line.EndArrowheadWidth = msoArrowheadWide
    '名前
    connectShape.Name = "connect-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
End Function
'コネクタの種類を直線、曲線、エルボーで切り替える
Sub AG_コネクタ種類切り替え()
Attribute AG_コネクタ種類切り替え.VB_ProcData.VB_Invoke_Func = "i\n14"

    If connectorErrorCheck Then
       Exit Sub
    End If
    
    With Selection.ShapeRange.ConnectorFormat
        If .Type = msoConnectorElbow Then
            .Type = msoConnectorCurve
        ElseIf .Type = msoConnectorCurve Then
            .Type = msoConnectorStraight
        Else
            .Type = msoConnectorElbow
        End If
    End With
End Sub
'選択したコネクタの始点側のシェイプとの接続位置を変更する
Sub AH_コネクタ始点変更()
Attribute AH_コネクタ始点変更.VB_ProcData.VB_Invoke_Func = "o\n14"

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
Sub AI_コネクタ終点変更()
Attribute AI_コネクタ終点変更.VB_ProcData.VB_Invoke_Func = "p\n14"
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
Sub AJ_選択中シェイプをグループ化()
    Selection.Group.Select
End Sub
'#選択中のシェイプをグループ解除
Sub AK_選択中シェイプをグループ解除()
    Selection.Ungroup
End Sub
'#選択中のシェイプを最背面にする
Sub AL_選択中シェイプを最背面へ()
    If TypeName(Selection) = "Range" Then
        MsgBox "シェイプを一つ選んでから実行してね。"
        Exit Sub
    End If
    
    For Each shp In Selection.ShapeRange
        shp.ZOrder msoSendToBack
    Next
End Sub
'#選択中のシェイプを最背面にする
Sub AM_選択中シェイプを最前面へ()
    If TypeName(Selection) = "Range" Then
        MsgBox "シェイプを一つ選んでから実行してね。"
        Exit Sub
    End If
    
    For Each shp In Selection.ShapeRange
        shp.ZOrder msoBringToFront
    Next
End Sub
'#シェイプの場所に値がなくなるように空行を挿入する
Sub AN_シェイプ貼付時ブランク行挿入()
    '#クリップボードにデータがある時のみ
    If Application.ClipboardFormats(1) Then
        '貼付。Ctrl + Vにあたるアクション（この時点でSelectionはシェイプになる）
        ActiveSheet.Paste
        
        '移動位置を取得するためのダミーシェイプ
        Dim dummyShape As Shape
        
        '左下隅のセルを取得するためのダミーシェイプ
        Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, Selection.top + Selection.Height, 1, 1)
    
        '「セルに合わせて移動やサイズ変更をしない」に設定
        'これやらないと行の挿入に合わせてシェイプも一緒に伸びてしまうから
        Selection.Placement = xlFreeFloating
        
        '貼付シェイプの下にあるセル分ループ
        For i = Selection.TopLeftCell.Row To dummyShape.TopLeftCell.Row + 1
            '列ループ（■j=Selection.TopLeftCell.Columnならシェイプ貼付位置から開始）
            For j = 1 To 15
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
Sub AO_目次シートを作成する()
    Dim ws As Worksheet
    
    '関数は別途参照
    If Not isExistCheckToSheet(ActiveWorkbook, "目次") Then
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
        ws.Cells(1, 9) = "更新者"
        ws.Cells(1, 10) = "更新日"
        'フォント色
        Range("A1:J1").Font.Color = RGB(20, 10, 10)
        '背景色
        Range("A1:J1").Interior.Color = RGB(255, 242, 204)
        '太字
        Range("A1:J1").Font.Bold = True
        Cells(2, 1).Select
        'ウィンドウ枠の固定
        ActiveWindow.FreezePanes = True
        '目盛線非表示
        ActiveWindow.DisplayGridlines = False
    Else
        Set ws = Sheets(1)
    End If
    
    Dim loopWs As Worksheet
    
    For i = 2 To Worksheets.count
        Set loopWs = Worksheets(i)
        ws.Cells(i, 1) = i - 1
        ws.Cells(i, 2) = loopWs.Name
        ws.Cells(i, 4) = loopWs.Shapes.count
        ws.Cells(i, 5) = loopWs.UsedRange.Address
        'シートではなくブック単位の情報なためコメントアウト
'        ws.Cells(i, 7) = ActiveWorkbook.BuiltinDocumentProperties(3)
'        ws.Cells(i, 8) = ActiveWorkbook.BuiltinDocumentProperties(11)
'        ws.Cells(i, 8).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
'        ws.Cells(i, 9) = ActiveWorkbook.BuiltinDocumentProperties(7)
'        ws.Cells(i, 10) = ActiveWorkbook.BuiltinDocumentProperties(12)
'        ws.Cells(i, 10).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    Next
    ws.Columns("A:H").AutoFit
    '■必要があれば以下をコメントイン
    ws.Cells(i + 1, 1) = "必要に応じて下記の関数を追加する。目次シートへのショートカット関数"
    ws.Cells(i + 2, 1) = "Sub 目次シートを選択"
    ws.Cells(i + 3, 1) = "    Sheets(1).Select"
    ws.Cells(i + 4, 1) = "End Sub"
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
        isExistCheckToSheet = checkSheet <= wb.Worksheets.count
    End If
End Function
'#アクティブシートの内容に従いシートを生成し、リンクを付与する。ソートも行う
Sub AP_シート生成とリンク付与()

    '[目次]シート上での実行を想定している。
    Dim topSheet As Worksheet
    Set topSheet = ActiveSheet
    
    '2列目に値のある最後の行を取得する
    Dim lastRowToBottom As Integer: lastRowToBottom = topSheet.Cells(1, 2).End(xlDown).Row
    
    Dim sheetName As String
    Dim linkRange As Range

    For i = 2 To lastRowToBottom
        sheetName = topSheet.Cells(i, 2).Value
        Set linkRange = topSheet.Cells(i, 2)
        
        If Not existsSheet(sheetName) Then
            'シートが存在していない場合
            With Worksheets.Add(after:=ActiveSheet)
                'シートを生成し、リンクを付与する
                .Name = sheetName
                topSheet.Hyperlinks.Add Anchor:=linkRange, Address:="", SubAddress:="'" & .Name & "'!A1"
                .Select
                '目盛線非表示、シート縮尺調整
                ActiveWindow.DisplayGridlines = False
                ActiveWindow.Zoom = 75
            End With
        Else
            '既にシートがある場合
            topSheet.Hyperlinks.Add Anchor:=linkRange, Address:="", SubAddress:="'" & sheetName & "'!A1"
            Sheets(sheetName).Select
                '目盛線非表示、シート縮尺調整
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = 75
            Sheets(sheetName).Cells(1, 1).Select
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
Sub AQ_シェイプ追加整列()
Attribute AQ_シェイプ追加整列.VB_ProcData.VB_Invoke_Func = " \n14"
    '■画像間の間隔
    Const MARGIN_BOTTOM = 70
    
    
    '貼付座標を格納する（topは都度書き換え、leftは初期値を使いまわす）
    Dim top As Integer: top = Selection.top + 5
    
    'キャプションを記載する用のセル
    Dim captionRange As Range
    Dim moveShape As Shape 'キャプションタイトル
    Dim captionText As String: captionText = "▼"
    For Each moveShape In ActiveSheet.Shapes
        '次に該当しないものは対象外：画像、グループ
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
        top = top + moveShape.Height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).Height
CONTINUE:
    Next
End Sub
'シェイプ整列2カラム編（奇数が1列目、偶数は2列目）
Sub AR_シェイプを2カラムで並べる()
    '■画像間の間隔
    Const MARGIN_BOTTOM = 70
    
    '貼付座標を格納する（topは都度書き換え、leftは初期値を使いまわす）
    Dim top As Integer: top = Selection.top + 5
    Dim left As Integer: left = Selection.left
    
    'キャプションを記載する用のセル
    Dim captionRange As Range
    Dim moveShape As Shape
    
    'エラーチェック
    If Selection.Row - REMARK_LINE - 1 < 1 Then
        MsgBox "キャプション用の行が足りません。あと" & REMARK_LINE - Selection.Row + 2 & "行下の位置で実行してください。"
        Exit Sub
    End If
    
    'キャプションタイトル
    Dim captionText As String
    '■ダイアログを使う場合は以下のコメントアウト分を使用する
    If CAPTION_TEXT_TOP_FLAG Then
        captionText = "▼変更前→変更後" 'InputBox("キャプションの初期値を入れて。", "キャプションオプション", "▼ここに画像の説明を書く")
    Else
        captionText = "▲変更前→変更後" 'InputBox("キャプションの初期値を入れて。", "キャプションオプション", "▲ここに画像の説明を書く")
    End If
    
    If StrPtr(answer) = 0 Then
        'キャンセル時
        Exit Sub
    End If
    Dim count As Integer: count = 1
    For Each moveShape In ActiveSheet.Shapes
        '次に該当しないものは対象外：画像、グループ、塗りつぶしのないオートシェイプ
        '条件変更時参考：https://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
        If moveShape.Type <> msoPicture _
            And moveShape.Type <> msoGroup _
            And (moveShape.Type = msoAutoShape) Then  'And Not moveShape.Fill.Visible
            GoTo CONTINUE:
        End If
        
        '移動位置を取得するためのダミーシェイプ
        Dim dummyShape As Shape
    
        '左上隅のセルを取得するためのダミーシェイプ
        Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
        
        'シェイプを移動する
        moveShape.top = dummyShape.TopLeftCell.Offset(0, 0).top
        If count Mod 2 = 1 Then
            '左右に並ぶシェイプの左側
            moveShape.left = Selection.left
        ElseIf count Mod 2 = 0 Then
            '左右に並ぶシェイプの右側
            moveShape.left = left
        End If
        
        moveShape.Name = "image-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
        
        'キャプション入力用セルを取得する（-1はタイトル分）
        Set captionRange = dummyShape.TopLeftCell.Offset(-1 - REMARK_LINE, 0)
        
        '用済みだから削除する
        dummyShape.Delete
        
        '■キャプション入力の設定（不要ならコメントアウトして）
        Call setCaption(captionRange, captionText)
        
        If count Mod 2 = 1 Then
            'topに与える数値を変えない。次のシェイプに与えるleftプロパティ値を設定する
            left = Selection.left + moveShape.WIDTH - 20
            '■変化を示す矢印。不要ならコメントアウトして。
            'Set onShape = ActiveSheet.Shapes.AddShape(msoShapeRightArrow, left - 10, top + moveShape.Height / 2, 40, 50)
        ElseIf count Mod 2 = 0 Then
            '今対象にしたシェイプの上部座標 + 今対象にしたシェイプの高さ + 画像間の間隔 + キャプションセル行の高さ = 次のシェイプの移動先上部座標
            top = top + moveShape.Height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).Height
        End If
        count = count + 1
CONTINUE:
    Next
    
    'END処理
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
    Call setCaption(dummyShape.TopLeftCell, "- END -")
    dummyShape.Delete
    
End Sub
'図形の矢印を付与する
Sub AS_シェイプ間に図形矢印を置く()
Attribute AS_シェイプ間に図形矢印を置く.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim startShape As Shape
    Dim endShape As Shape
    Dim connectShape As Shape
    
    Dim x1 As Double
    Dim x2 As Double
    Dim y1 As Double
    Dim y2 As Double
    Dim degree As Double
    Dim adjustDegree As Integer
    Dim onShape As Shape
    Dim setLeft As Double
    Dim setTop As Double
    
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
    
    For i = 1 To Selection.ShapeRange.count - 1
        '選択中シェイプの保持（元）
        Set startShape = Selection.ShapeRange.Item(i)
        '選択中シェイプの保持（先）
        Set endShape = Selection.ShapeRange.Item(i + 1)
        '各ポイントを取得
        x1 = startShape.left + startShape.WIDTH - (startShape.WIDTH / 2)
        x2 = endShape.left + endShape.WIDTH - (endShape.WIDTH / 2)
        y1 = startShape.top + startShape.Height - (startShape.Height / 2)
        y2 = endShape.top + endShape.Height - (endShape.Height / 2)
        
        '左位置プロパティの調整
        If startShape.left < endShape.left Then
            adjustDegree = 180
            setLeft = startShape.left + startShape.WIDTH + ((endShape.left - (startShape.left + startShape.WIDTH)) / 2) - 25
        Else
            setLeft = endShape.left + endShape.WIDTH + ((startShape.left - (endShape.left + endShape.WIDTH)) / 2) - 25
            
        End If
        
        '上位置プロパティの調整
        If startShape.top < endShape.top Then
            setTop = startShape.top + (startShape.Height / 2) + (((endShape.top + (endShape.Height / 2)) - (startShape.top + (startShape.Height / 2))) / 2) - 25
        Else
            setTop = endShape.top + (endShape.Height / 2) + (((startShape.top + (startShape.Height / 2)) - (endShape.top + (endShape.Height / 2))) / 2) - 25
        End If
        '矢印の向き調整。最後に割ってるのは円周率
        If x2 - x1 <> 0 Then
            degree = Atn((y2 - y1) / (x2 - x1)) * 180 / 3.14
        Else
            degree = -90
        End If
        
        Set onShape = ActiveSheet.Shapes.AddShape(msoShapeLeftArrow, setLeft, setTop, 50, 50)
        onShape.Name = "allow-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
        onShape.Rotation = degree + adjustDegree
        adjustDegree = 0
    Next
End Sub
'セルからセルへ枠シェイプを繋ぐ矢印を付与する
Sub AT_セルからセルに伸びるコネクタ_枠あり()
    Dim onShape As Shape
    
    For Each rcell In Selection
        Set onShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, _
                                                    rcell.MergeArea.left, _
                                                    rcell.MergeArea.top, _
                                                    rcell.MergeArea.WIDTH, _
                                                    rcell.MergeArea.Height)
        onShape.Name = "shape-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
        '■塗りつぶし（msoTrue:あり、msoFalse:なし）
        onShape.Fill.Visible = msoFalse
        '■線の太さ。お好みでどうぞ
        onShape.Line.Weight = 1
        '■色指定
        onShape.Line.ForeColor.RGB = RGB(0, 0, 0)
        '■塗りつぶし色指定
        'onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
        'コネクタで繋ぐため、選択状態にする
        onShape.Select Replace:=False
    Next
    'シェイプ間に矢印を付与していく
    Call AF_シェイプを選択順にコネクタで繋ぐ
End Sub
'色やスタイルを変更したいとき、一つだけに実施し、あとはコピーするというもの
Sub AU_最初のシェイプをスキャンコピー()
    '複数シェイプ選択時、2つ目以降は1つ目のスタイルを適用する。ループ処理
    Dim baseShp As Shape
    Dim shp As Shape
    Set baseShp = Selection.ShapeRange.Item(1)
    For i = 2 To Selection.ShapeRange.count
        '選択中シェイプの保持（接続元）
        Set shp = Selection.ShapeRange.Item(i)
        shp.Line.ForeColor.RGB = baseShp.Line.ForeColor.RGB
        'shp.Line.Weight = baseShp.Line.Weight
        'shp.ForeColor.RGB = baseShp.ForeColor.RGB
        '■↓ワードアートフォーマットが指定できないシェイプを選ぶときはコメントアウトして
        'shp.TextFrame2.WordArtformat = baseShp.TextFrame2.WordArtformat
        shp.Fill.Transparency = baseShp.Fill.Transparency
        '■テキストまで変えたくないときはコメントアウト
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
        '■なぜか実行時にエラー
        'shp.Shadow.Type = baseShp.Shadow.Type
        shp.Shadow.Visible = baseShp.Shadow.Visible
        '■なぜか実行時にエラー
        'shp.Shadow.Style = baseShp.Shadow.Style
        shp.Shadow.Blur = baseShp.Shadow.Blur
        shp.Shadow.OffsetX = baseShp.Shadow.OffsetX
        shp.Shadow.OffsetY = baseShp.Shadow.OffsetY
        shp.Shadow.RotateWithShape = baseShp.Shadow.RotateWithShape
        shp.Shadow.ForeColor.RGB = baseShp.Shadow.ForeColor.RGB
        '■なぜか実行時にエラー
        'shp.Shadow.Transparency = baseShp.Shadow.Transparency
        shp.Shadow.Size = baseShp.Shadow.Size
    Next
End Sub
'省略線をだす。黒、白、黒の3本のにょろにょろ線を作成し、最後にグループ化している
Sub AV_省略にょろにょろ出現()
    Dim selectRange As Range
    Set selectRange = Selection
    Dim top As Integer: top = Selection.top
    Dim left As Integer: left = Selection.left
    '■にょろにょろの長さを設定する。数字が大きいと結構時間がかかる
    Const WIDTH = 50
    Dim blackTopShape As Shape
    Dim whiteShape As Shape
    Dim blackBottomShape As Shape

    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, left, top)
        For i = 0 To WIDTH
            If i Mod 2 = 0 Then
                top = top + 5
            Else
                top = top - 5
            End If
            left = left + 7
            .AddNodes msoSegmentCurve, msoEditingAuto, left, top
        Next
        Set blackTopShape = .ConvertToShape
    End With
    blackTopShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    blackTopShape.Line.Weight = 3
    blackTopShape.Name = "omit-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()

    top = top + 3
    left = Selection.left
    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, left, top)
    
        For i = 0 To WIDTH
            If i Mod 2 = 0 Then
                top = top + 5
            Else
                top = top - 5
            End If
            left = left + 7
            .AddNodes msoSegmentCurve, msoEditingAuto, left, top
        Next
        Set blackBottomShape = .ConvertToShape
    End With
    blackBottomShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    blackBottomShape.Line.Weight = 3
    blackBottomShape.Name = "omit-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
    
    
    top = top - 9
    left = Selection.left
    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, left, top)
    
        For i = 0 To WIDTH
            If i Mod 2 = 0 Then
                top = top + 5.2
            Else
                top = top - 5.2
            End If
            left = left + 7
            .AddNodes msoSegmentCurve, msoEditingAuto, left, top
        Next
        Set whiteShape = .ConvertToShape
    End With
    whiteShape.Line.ForeColor.RGB = RGB(255, 255, 255)
    whiteShape.Line.Weight = 5.8
    whiteShape.Name = "omit-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
    
    blackTopShape.Select Replace:=False
    whiteShape.Select Replace:=False
    blackBottomShape.Select Replace:=False
    Selection.Group
    selectRange.Select
End Sub
'色を確認しますねん
Sub AW_色確認()
Attribute AW_色確認.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim fontColorCode As Long
    Dim backColorCode As Long
    Dim lineColorCode As Long

    If TypeName(Selection) = "Range" Then
        Debug.Print "▼セルの情報（" & Now() & "）──────-┐"
        'セルの背景色
        backColorCode = Selection.Interior.Color
        'セルの文字色
        fontColorCode = Selection.Font.Color
        'セルの枠線色
        lineColorCode = Selection.Borders.Color
        
    ElseIf Selection.ShapeRange.Connector Then
        Debug.Print "▼矢印の情報（" & Now() & "）──────-┐"
        '矢印の枠線色
        lineColorCode = Selection.ShapeRange.Item(1).Line.ForeColor.RGB
        
    ElseIf TypeName(Selection) = "Rectangle" Then
        Debug.Print "▼シェイプの情報（" & Now() & "）────-┐"
        'シェイプの塗りつぶし色
        backColorCode = Selection.ShapeRange.Item(1).Fill.ForeColor.RGB
        'シェイプの文字色
        fontColorCode = Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
        'シェイプの枠線色
        lineColorCode = Selection.ShapeRange.Item(1).Line.ForeColor.RGB
    Else
        Debug.Print "▼ここに入った場合想定外"
    End If

    Debug.Print "│ 背景色/塗りつぶし色の色コード：" & backColorCode
    Dim Red As Integer: Red = backColorCode Mod 256
    Dim Green As Integer: Green = Int(backColorCode / 256) Mod 256
    Dim Blue As Integer: Blue = Int(backColorCode / 256 / 256)
    Debug.Print "│ 背景色/塗りつぶし色のRGB値：RGB(" & Red & "," & Green; "," & Blue & ")"
    Debug.Print "├───────────────────────┤"
    Debug.Print "│ 文字色の色コード：" & fontColorCode
    Red = fontColorCode Mod 256
    Green = Int(fontColorCode / 256) Mod 256
    Blue = Int(fontColorCode / 256 / 256)
    Debug.Print "│ 文字色のRGB値：RGB(" & Red & "," & Green; "," & Blue & ")"
    Debug.Print "├───────────────────────┤"
    Debug.Print "│ セル枠色/線色の色コード：" & lineColorCode
    Red = lineColorCode Mod 256
    Green = Int(lineColorCode / 256) Mod 256
    Blue = Int(lineColorCode / 256 / 256)
    Debug.Print "│ セル枠色/線色のRGB値：RGB(" & Red & "," & Green; "," & Blue & ")"
    Debug.Print "└───────────────────────┘"

End Sub
Sub AX_X座標合わせ()
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).left + Selection.ShapeRange.Item(1).WIDTH / 2
    For Each sp In Selection.ShapeRange
        sp.left = baseXCenterPoint - sp.WIDTH / 2
    Next
End Sub
Sub AY_Y座標合わせ()
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).top + Selection.ShapeRange.Item(1).Height / 2
    For Each sp In Selection.ShapeRange
        sp.top = baseXCenterPoint - sp.Height / 2
    Next
End Sub
'セルからセルへ矢印をつける。1→2,3→4のようにつける
Sub AZ_セルからセルに伸びるコネクタ_枠なし()
    Dim count As Integer: count = 1
    '始点セル
    Dim startRange As Range
    '終点セル
    Dim endRange As Range
    '矢印シェイプ
    Dim connectShape As Shape
    For Each cell In Selection
        If count Mod 2 = 1 Then
            '奇数の時は始点セルを変数へ
            Set startRange = cell
        ElseIf count Mod 2 = 0 Then
            '偶数の時は終点セルの設定と矢印の挿入
            Set endRange = cell
            Set connectShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorElbow, _
                BeginX:=startRange.left + (startRange.WIDTH), BeginY:=startRange.top + (startRange.Height / 2), _
                EndX:=endRange.left + (endRange.WIDTH / 2), EndY:=endRange.top + (endRange.Height / 2))
            Call makeConnectAllow(connectShape)
        End If
        count = count + 1
    Next
End Sub
'指定色内で色の切り替えを行う。必要に応じてコメントアウトの位置、色設定プロパティを調整して
Sub BA_色切り替え()
Attribute BA_色切り替え.VB_ProcData.VB_Invoke_Func = " \n14"
    'シェイプの塗りつぶし色
    'Dim colorCode As String: colorCode = Selection.ShapeRange.Item(1).Fill.ForeColor.RGB
    'シェイプの文字色
    'Dim colorCode As String: colorCode = Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    'シェイプの枠線色
    Dim colorCode As String: colorCode = Selection.ShapeRange.Item(1).Line.ForeColor.RGB

    Dim colors() As Variant
    '■色コードリスト（「AW_色確認」でコードを取得し設定する）
    colors = Array("255", "49407", "65535", "5296274", "5287936", "15773696", "12611584")
    
    Dim hitIndex As Integer
    hitIndex = isExistArrayReturnIndex(colors, colorCode)
    
    If hitIndex <> -1 Then
        If UBound(colors) = hitIndex Then
            '色配列の最終インデックスだった場合、0番目に戻す
            Selection.ShapeRange.Item(1).Line.ForeColor.RGB = colors(0)
        Else
            '色配列に属しており、次のインデックス色に設定する
            Selection.ShapeRange.Item(1).Line.ForeColor.RGB = colors(hitIndex + 1)
        End If
    Else
        '色配列のどれにもあたらなかった場合
        Selection.ShapeRange.Item(1).Line.ForeColor.RGB = colors(0)
    End If
    
End Sub
