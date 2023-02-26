Attribute VB_Name = "main"
'ツール＞参照設定から「Microsoft XML, v6.0」をONにする
Dim ws As Worksheet
'ドキュメントオブジェクト
Dim doc As MSXML2.DOMDocument60
'走査時にXML要素オブジェクトを格納する
Dim XMLchild As Object 'ループ中の処理対象によって、入るオブジェクトが異なるから、詳細なオブジェクトは指定しないでおく。
'書き出し行
Dim writeLine As Integer
'走査中要素の階層レベル
Dim level As Integer
'属性記録開始列（処理中に属性数に応じてインクリメントする）
Dim attributesWriteCol As Integer
Const FILE_PATH = "C:\Users\tp04372\Documents\macro\研究所\sample.xml"
Sub trial2()
    Set ws = ActiveSheet
    ws.Cells.Clear
    ws.Cells(1, 1) = "兄弟ノード数"
    ws.Cells(1, 2) = "子ノード数"
    ws.Cells(1, 3) = "親要素名"
    ws.Cells(1, 4) = "レベル"
    ws.Cells(1, 5) = "要素名"
    ws.Cells(1, 6) = "要素内容"
    ws.Cells(1, 7) = "要素タイプ"
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    '対象ファイル読み込み
    doc.Load FILE_PATH
    'それぞれ初期化
    writeLine = 1
    level = 0
    attributesWriteCol = 7
    
    Call writeXMLElement(doc)
    '初期化。
    Set doc = Nothing
End Sub
'#渡されたXML親要素について再帰的に子要素を呼び出しその内容をセルに書き出す
Private Sub writeXMLElement(XMLparent As Variant)
'    Debug.Print XMLparent.ChildNodes.Length
    
    'Debug.Print XMLparent.parseError.reason
    
    If XMLparent.ChildNodes.Length <> 0 Then
        For Each XMLchild In XMLparent.ChildNodes
        
'            Debug.Print XMLparent.ChildNodes.Length
'            Debug.Print XMLchild.ChildNodes.Length
            '要素内容の場合、子ノード数が0になる。出力の意味がないからスキップさせる
            If XMLparent.ChildNodes.Length <> 0 Then
                level = level + 1
                writeLine = writeLine + 1
                ws.Cells(writeLine, 1) = XMLparent.ChildNodes.Length 'これの意味がいまいちわかっていない
                ws.Cells(writeLine, 2) = XMLchild.ChildNodes.Length
                ws.Cells(writeLine, 3) = XMLchild.BaseName
                ws.Cells(writeLine, 4) = level
                ws.Cells(writeLine, 5) = XMLchild.nodeName
                ws.Cells(writeLine, 6) = IIf(XMLparent.ChildNodes.Length <> 1, "", XMLchild.Text)
                ws.Cells(writeLine, 7) = XMLchild.nodeTypeString

                '属性出力
                For Each memberAttribute In XMLchild.Attributes
                    attributesWriteCol = attributesWriteCol + 1
                    ws.Cells(writeLine, attributesWriteCol) = memberAttribute.Name & "：" & memberAttribute.Value
                Next
                '属性の数によってインクリメントした分を初期化する
                attributesWriteCol = 7
                '==xmlファイルの高さを再現して出すソース===========
                'ここをコメントインする場合、通常出力部分はコメントアウトする。なお属性出力は対象外
'                ws.Cells(writeLine, level) = XMLchild.nodeName & IIf(XMLparent.ChildNodes.Length <> 1, "", "：" & XMLchild.Text)
                '==================================================
                Call writeXMLElement(XMLchild)
                '再帰処理から一段元に戻るため階層レベルも一つ戻す
                level = level - 1
            End If
        Next
    End If
End Sub
