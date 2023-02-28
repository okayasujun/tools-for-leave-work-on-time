Attribute VB_Name = "try"
'#シェイプを整列させる（Y座標の昇順）
'※既知の問題：同じ高さの画像があるとエラーを出す。
Sub E_高さ順に整列させる() 'TODO:これに関してはシンプルにリファクタリングをすべし。共通部分とか、きれいにしよう。
Attribute E_高さ順に整列させる.VB_ProcData.VB_Invoke_Func = "t\n14"

    '■画像間の間隔
    Const MARGIN_BOTTOM = 70

    '画像の数を取得（https://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype）
    Dim pictureCount As Integer: pictureCount = 0
    For Each moveShape In ActiveSheet.Shapes
        If moveShape.Type = msoPicture Or moveShape.Type = msoGroup Then
            pictureCount = pictureCount + 1
        End If
    Next

    '画像名とY軸上部座標をマップで保持
    'ツール＞参照設定から「Microsoft Scripting Runtime」の有効化が必要
    Dim shapeDic As Dictionary
    Set shapeDic = CreateObject("Scripting.Dictionary")
    
    Dim Count As Integer: Count = 0
    Dim shapeYArray() As Double
    
    'マップに情報を設定、座標ソート用に座標情報の配列格納
    For Each moveShape In ActiveSheet.Shapes
        If moveShape.Type = msoPicture Or moveShape.Type = msoGroup Then
        
            If Count = 0 Then
                ReDim shapeYArray(Count)
                shapeYArray(Count) = moveShape.top
            Else
                ReDim Preserve shapeYArray(Count)
                shapeYArray(Count) = moveShape.top
            End If
            
            shapeDic.Add moveShape.Name, moveShape.top
            Count = Count + 1
        End If
    Next
    
    'topプロパティの順にソート
    shapeYArray = sort(shapeYArray)
    
    '後続処理において同じ高さの画像がある場合、キー重複エラーになるのでcatch準備
    On Error GoTo ErrHndl
    
    'ソート後のシェイプ情報を詰めなおす
    Dim sortedShapeDic As Dictionary
    Set sortedShapeDic = CreateObject("Scripting.Dictionary")

    For Each yPoint In shapeYArray
        For Each dKey In shapeDic
            If yPoint = shapeDic.Item(dKey) Then
                sortedShapeDic.Add dKey, shapeDic.Item(dKey)
                Exit For
            End If
        Next
    Next

    '============================== 以降の処理はlineUpShapesOrderOfPasted()と同じ =================================

    '移動位置を取得するためのダミーシェイプ
    Dim dummyShape As shape
    '貼付座標を格納する（topは都度書き換え、leftは初期値を使いまわす）
    Dim top As Integer: top = Selection.top + 5
    Dim left As Integer: left = Selection.left
    'キャプションを記載する用のセル
    Dim captionRange As Range

    For Each dKey In sortedShapeDic
    
        Set moveShape = ActiveSheet.Shapes(dKey)
        
        '左上隅のセルを取得するためのダミーシェイプ
        Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, left, top, 1, 1)
        
        'シェイプを移動する
        moveShape.top = dummyShape.TopLeftCell.top
        moveShape.left = left
        
        'キャプション入力用セルを取得する
        Set captionRange = dummyShape.TopLeftCell
        
        '用済みだから削除する
        dummyShape.Delete
        
        '■キャプション入力用
        'Call setCaption(captionRange)
        
        '今対象にしたシェイプの上部座標 + 今対象にしたシェイプの高さ + 画像間の間隔 = 次のシェイプの移動先上部座標
        top = top + moveShape.height + MARGIN_BOTTOM
    Next
    Exit Sub
    
ErrHndl:
    MsgBox "同じ高さの画像があるから、少しずらしてリトライしてね。"
End Sub
'ソート開始
Function sort(ByRef targetArray() As Double)
    Dim swap As Double
    'ソート開始
    For i = LBound(targetArray) To UBound(targetArray)
        For j = UBound(targetArray) To i Step -1
            If targetArray(i) > targetArray(j) Then
                swap = targetArray(i)
                targetArray(i) = targetArray(j)
                targetArray(j) = swap
            End If
        Next j
    Next i
    sort = targetArray
End Function
'============================================================================================================================
'#TODO:必要性・使い勝手の観点で要見直し。グループ化対象範囲をシェイプ内とするか選択セル範囲内とするか。
'#選択中の大枠内にあるシェイプをグループ化する。グループ完了時、囲みシェイプは削除する
Sub M_選択中枠内のシェイプ群をグループ化する()
    'グループ化シェイプ名をカンマ区切りで保持する用
    Dim targetShapeName As String
    'カンマ区切りで保持したものを配列状態で保持するよう
    Dim targetShapeArray As Variant
    
    For Each shape In ActiveSheet.Shapes
        '条件参考：https://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype

        If shape.Type = msoAutoShape Or shape.Type = msoGroup Or shape.Type = msoPicture Then
            '上辺、左辺、右辺、下辺が大枠内にあるシェイプのみを対象とする
            If Selection.left < shape.left _
                And Selection.top < shape.top _
                And shape.left + shape.width < Selection.left + Selection.width _
                And shape.top + shape.height < Selection.top + Selection.height Then
                'グループ対象シェイプの記録（後続処理でグループ化）
                targetShapeName = targetShapeName & shape.Name & ","
            End If
        End If
        
    Next
    '対象シェイプを囲っていたシェイプを削除する
    Selection.Delete
    'グループ対象シェイプ名を配列化
    targetShapeArray = Split(targetShapeName, ",")
    
    For Each shape In ActiveSheet.Shapes
        '全シェイプの中からグループ対象のものだけ選択状態にする
        If isExistArray(targetShapeArray, shape.Name) Then
            shape.Select Replace:=False
        End If
    Next
    
    On Error GoTo catch
    
    If VarType(Selection) = vbObject Then
        '選択中シェイプをグループ化
        Selection.Group.Select
    End If
    
    Exit Sub
catch:
End Sub
'#配列内に存在するかどうか
Function isExistArray(targetArray As Variant, checkValue As String)
    isExistArray = False
    
    If UBound(targetArray) = -1 Then
        'UBoundの戻り値：-1は要素数0を示す。この場合、すべて対象外とする
        isExistArray = False
        Exit Function
    End If
    
    For i = LBound(targetArray) To UBound(targetArray)
        If targetArray(i) = checkValue Then
            isExistArray = True
            Exit For
        End If
    Next
End Function
'============================================================================================================================
'作りかけ20230228。ブラッシュアップの余地あり
Sub S_シェイプ間にマニュアル向けの矢印を置く()
    Dim startShape As shape
    Dim endShape As shape
    Dim connectShape As shape
    
    Dim x1 As Double
    Dim x2 As Double
    Dim y1 As Double
    Dim y2 As Double
    Dim degree As Double
    Dim adjustDegree As Integer
    Dim onShape As shape
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
    
    For i = 1 To Selection.ShapeRange.Count - 1
        '選択中シェイプの保持（接続元）
        Set startShape = Selection.ShapeRange.Item(i)
        '選択中シェイプの保持（接続先）
        Set endShape = Selection.ShapeRange.Item(i + 1)
        '各ポイントを取得
        x1 = startShape.left + startShape.width - (startShape.width / 2)
        x2 = endShape.left + endShape.width - (endShape.width / 2)
        y1 = startShape.top + startShape.height - (startShape.height / 2)
        y2 = endShape.top + endShape.height - (endShape.height / 2)
    
        If startShape.left < endShape.left Then
            adjustDegree = 180
            setLeft = startShape.left + (endShape.left - startShape.left) / 2
            setTop = startShape.top + (endShape.top - startShape.top) / 2
        Else
            setLeft = endShape.left + (startShape.left - endShape.left) / 2
            setTop = endShape.top + (startShape.top - endShape.top) / 2
        
        End If
        
        degree = Atn((y2 - y1) / (x2 - x1)) * 180 / 3.14
        
        Set onShape = ActiveSheet.Shapes.AddShape(msoShapeLeftArrow, setLeft, setTop, 200, 100)
        onShape.Rotation = degree + adjustDegree
        adjustDegree = 0
    Next
    '状況：やりたいことのイメージはOK
    'あとは矢印のスタイル、大きさ、位置の微調整
End Sub
'============================================================================================================================
'謝意：https://www.ka-net.org/blog/?p=4944 参考
'できたけど動作不安定（クリップボードの表示エリア中可視範囲のものしか対象にできない）
Sub T_連続貼付()
    'TODO:実行前後でbeforeセル選ばせる処理入れてもいいかも。今、最後の貼付シェイプを選んだ状態になる
    'Officeクリップボードにあるアイテム列挙
    Dim aryListItems As UIAutomationClient.IUIAutomationElementArray
    Dim i As Long
    Dim ptnAcc As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
   
    Set aryListItems = GetOfficeClipboardListItems
    For i = 0 To aryListItems.Length - 1
        Debug.Print i + 1, aryListItems.GetElement(i).CurrentName
    
        '=============
        Set ptnAcc = aryListItems.GetElement(i).GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        ptnAcc.DoDefaultAction
    Next
    'ここでクリップボードの表示をfalseに戻してはだめ
End Sub
'
Sub U_クリップボードすべてクリア()
    DoActionOfficeClipboard "すべてクリア"
End Sub
'ボタン操作を実行する（「すべてクリア」でのみ使用する）
Private Sub DoActionOfficeClipboard(ByVal ButtonName As String)
'Officeクリップボードコマンド実行
  Dim uiAuto As UIAutomationClient.CUIAutomation
  Dim accClipboard As Office.IAccessible
  Dim elmClipboard As UIAutomationClient.IUIAutomationElement
  Dim elmButton As UIAutomationClient.IUIAutomationElement
  Dim cndButtons As UIAutomationClient.IUIAutomationCondition
  Dim aryButtons As UIAutomationClient.IUIAutomationElementArray
  Dim ptnAcc As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
  Dim i As Long
   
  Set elmButton = Nothing '初期化
  Set uiAuto = New UIAutomationClient.CUIAutomation
  With Application
    .CommandBars("Office Clipboard").Visible = True
    DoEvents
    Set accClipboard = .CommandBars("Office Clipboard")
  End With
  Set elmClipboard = uiAuto.ElementFromIAccessible(accClipboard, 0)
  Set cndButtons = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ButtonControlTypeId)
  Set aryButtons = elmClipboard.FindAll(TreeScope_Subtree, cndButtons)
  For i = 0 To aryButtons.Length - 1
    If aryButtons.GetElement(i).CurrentName = ButtonName Then
      Set elmButton = aryButtons.GetElement(i)
      Exit For
    End If
  Next
  If elmButton Is Nothing Then Exit Sub
  If elmButton.CurrentIsEnabled <> False Then
    Set ptnAcc = elmButton.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    ptnAcc.DoDefaultAction
  End If
End Sub
 
Private Function GetOfficeClipboardListItems() As UIAutomationClient.IUIAutomationElementArray
'Officeクリップボードリスト取得
  Dim uiAuto As UIAutomationClient.CUIAutomation
  Dim accClipboard As Office.IAccessible
  Dim elmClipboard As UIAutomationClient.IUIAutomationElement
  Dim cndListItems As UIAutomationClient.IUIAutomationCondition
   
  Set uiAuto = New UIAutomationClient.CUIAutomation
  With Application
    .CommandBars("Office Clipboard").Visible = True 'Falseにしてはだめ。
    DoEvents
    Set accClipboard = .CommandBars("Office Clipboard")
  End With
  Set elmClipboard = uiAuto.ElementFromIAccessible(accClipboard, 0)
  Set cndListItems = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ListItemControlTypeId)
  Set GetOfficeClipboardListItems = elmClipboard.FindAll(TreeScope_Subtree, cndListItems)
End Function
'============================================================================================================================
