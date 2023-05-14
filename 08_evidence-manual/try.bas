Attribute VB_Name = "try"
'#シェイプを整列させる（Y座標の昇順）
'※既知の問題：同じ高さの画像があるとエラーを出す。
Sub AE_シェイプを高さ順に整列させる() 'TODO:リファクタリングをすべし。共通部分とか、きれいにしよう。
Attribute AE_シェイプを高さ順に整列させる.VB_ProcData.VB_Invoke_Func = "t\n14"

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
    
    Dim count As Integer: count = 0
    Dim shapeYArray() As Double
    
    'マップに情報を設定、座標ソート用に座標情報の配列格納
    For Each moveShape In ActiveSheet.Shapes
        If moveShape.Type = msoPicture Or moveShape.Type = msoGroup Then
        
            If count = 0 Then
                ReDim shapeYArray(count)
                shapeYArray(count) = moveShape.top
            Else
                ReDim Preserve shapeYArray(count)
                shapeYArray(count) = moveShape.top
            End If
            
            shapeDic.Add moveShape.Name, moveShape.top
            count = count + 1
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
    Dim dummyShape As Shape
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
        Call setCaption(captionRange, "")
        
        '今対象にしたシェイプの上部座標 + 今対象にしたシェイプの高さ + 画像間の間隔 = 次のシェイプの移動先上部座標
        top = top + moveShape.Height + MARGIN_BOTTOM
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
''============================================================================================================================
'#TODO:必要性・使い勝手の観点で要見直し。グループ化対象範囲をシェイプ内とするか選択セル範囲内とするか。
'#選択中の大枠内にあるシェイプをグループ化する。グループ完了時、囲みシェイプは削除する
Sub YA_選択中枠内のシェイプ群をグループ化する()
    'グループ化シェイプ名をカンマ区切りで保持する用
    Dim targetShapeName As String
    'カンマ区切りで保持したものを配列状態で保持するよう
    Dim targetShapeArray As Variant

    For Each Shape In ActiveSheet.Shapes
        '条件参考：https://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype

        If Shape.Type = msoAutoShape Or Shape.Type = msoGroup Or Shape.Type = msoPicture Then
            '上辺、左辺、右辺、下辺が大枠内にあるシェイプのみを対象とする
            If Selection.left < Shape.left _
                And Selection.top < Shape.top _
                And Shape.left + Shape.WIDTH < Selection.left + Selection.WIDTH _
                And Shape.top + Shape.Height < Selection.top + Selection.Height Then
                'グループ対象シェイプの記録（後続処理でグループ化）
                targetShapeName = targetShapeName & Shape.Name & ","
            End If
        End If

    Next
    '対象シェイプを囲っていたシェイプを削除する
    Selection.Delete
    'グループ対象シェイプ名を配列化
    targetShapeArray = Split(targetShapeName, ",")

    For Each Shape In ActiveSheet.Shapes
        '全シェイプの中からグループ対象のものだけ選択状態にする
        If isExistArray(targetShapeArray, Shape.Name) Then
            Shape.Select Replace:=False
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
'指定された値が指定された配列内の何番目に存在するかを返す
Function isExistArrayReturnIndex(targetArray As Variant, checkValue As String)
    isExistArrayReturnIndex = -1

    If UBound(targetArray) = -1 Then
        'UBoundの戻り値：-1は要素数0を示す。この場合、-1を返す
        isExistArrayReturnIndex = -1
        Exit Function
    End If

    For i = LBound(targetArray) To UBound(targetArray)
        If targetArray(i) = checkValue Then
            isExistArrayReturnIndex = i
            Exit For
        End If
    Next
End Function
''============================================================================================================================
''謝意：https://www.ka-net.org/blog/?p=4944 参考
''できたけど動作不安定（クリップボードの表示エリア中可視範囲のものしか対象にできない）
'Sub YB_連続貼付()
'    'TODO:実行前後でbeforeセル選ばせる処理入れてもいいかも。今、最後の貼付シェイプを選んだ状態になる
'    'Officeクリップボードにあるアイテム列挙
'    Dim aryListItems As UIAutomationClient.IUIAutomationElementArray
'    Dim i As Long
'    Dim ptnAcc As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
'
'    Set aryListItems = GetOfficeClipboardListItems
'    For i = 0 To aryListItems.Length - 1
'        'Debug.Print i + 1, aryListItems.GetElement(i).CurrentName
'
'        '=============
'        Set ptnAcc = aryListItems.GetElement(i).GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
'        ptnAcc.DoDefaultAction
'    Next
'    'ここでクリップボードの表示をfalseに戻してはだめ
'End Sub
''
'Sub YC_クリップボードすべてクリア()
'    DoActionOfficeClipboard "すべてクリア"
'End Sub
''ボタン操作を実行する（「すべてクリア」でのみ使用する）
'Private Sub DoActionOfficeClipboard(ByVal ButtonName As String)
''Officeクリップボードコマンド実行
'  Dim uiAuto As UIAutomationClient.CUIAutomation
'  Dim accClipboard As Office.IAccessible
'  Dim elmClipboard As UIAutomationClient.IUIAutomationElement
'  Dim elmButton As UIAutomationClient.IUIAutomationElement
'  Dim cndButtons As UIAutomationClient.IUIAutomationCondition
'  Dim aryButtons As UIAutomationClient.IUIAutomationElementArray
'  Dim ptnAcc As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
'  Dim i As Long
'
'  Set elmButton = Nothing '初期化
'  Set uiAuto = New UIAutomationClient.CUIAutomation
'  With Application
'    .CommandBars("Office Clipboard").Visible = True
'    DoEvents
'    Set accClipboard = .CommandBars("Office Clipboard")
'  End With
'  Set elmClipboard = uiAuto.ElementFromIAccessible(accClipboard, 0)
'  Set cndButtons = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ButtonControlTypeId)
'  Set aryButtons = elmClipboard.FindAll(TreeScope_Subtree, cndButtons)
'  For i = 0 To aryButtons.Length - 1
'    If aryButtons.GetElement(i).CurrentName = ButtonName Then
'      Set elmButton = aryButtons.GetElement(i)
'      Exit For
'    End If
'  Next
'  If elmButton Is Nothing Then Exit Sub
'  If elmButton.CurrentIsEnabled <> False Then
'    Set ptnAcc = elmButton.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
'    ptnAcc.DoDefaultAction
'  End If
'End Sub
'
'Private Function GetOfficeClipboardListItems() As UIAutomationClient.IUIAutomationElementArray
''Officeクリップボードリスト取得
'  Dim uiAuto As UIAutomationClient.CUIAutomation
'  Dim accClipboard As Office.IAccessible
'  Dim elmClipboard As UIAutomationClient.IUIAutomationElement
'  Dim cndListItems As UIAutomationClient.IUIAutomationCondition
'
'  Set uiAuto = New UIAutomationClient.CUIAutomation
'  With Application
'    .CommandBars("Office Clipboard").Visible = True 'Falseにしてはだめ。
'    DoEvents
'    Set accClipboard = .CommandBars("Office Clipboard")
'  End With
'  Set elmClipboard = uiAuto.ElementFromIAccessible(accClipboard, 0)
'  Set cndListItems = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ListItemControlTypeId)
'  Set GetOfficeClipboardListItems = elmClipboard.FindAll(TreeScope_Subtree, cndListItems)
'End Function
''============================================================================================================================
