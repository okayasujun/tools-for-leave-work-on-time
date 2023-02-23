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
'End Sub
