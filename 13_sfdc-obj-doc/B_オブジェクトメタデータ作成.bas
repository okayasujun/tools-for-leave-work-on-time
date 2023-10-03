Attribute VB_Name = "B_オブジェクトメタデータ作成"
'
'objectのメタデータファイルを作成する
'[CustomObject]シートに定義した内容をベースに要所を置換する
'
'
Sub B_オブジェクトメタデータ作成()
    Dim objSheet As Worksheet
    Set objSheet = Sheets(OBJECT_SHEET)
    Dim objMetaSheet As Worksheet
    Set objMetaSheet = Sheets(OBJECT_META_SHEET)
    
    'オブジェクトのAPI名
    Dim objApiName As String: objApiName = Sheets(OBJECT_SHEET).Cells(4, 4).Value
    '生成対象ファイル名。例「CustomObject04__c.object-meta.xml」
    Dim fileName As String: fileName = ThisWorkbook.Path & "\objects\" & objApiName & "\" & objApiName & ".object-meta.xml"

    '連想配列作成（A,D列を対応付ける）
    Dim objectInformation As Object
    Set objectInformation = CreateObject("Scripting.Dictionary")
    For i = 1 To objSheet.Cells(Rows.Count, 1).End(xlUp).row
        With objSheet
            objectInformation(.Cells(i, 1).Value) = .Cells(i, 4).Value
        End With
    Next
    
    '正規表現（置換のため）
    Dim regexpObj As Object
    Set regexpObj = CreateObject("VBScript.RegExp")
    With regexpObj
        '置換文字抽出用パターン（VBAで肯定先読みは使えない）
        .Pattern = "{.*(?=})"
        '英大文字小文字を区別しない
        .IgnoreCase = True
        '文字列全体に対してパターンマッチさせる
        .Global = True
    End With
    'ファイルに書き出すテキスト
    Dim writeText As String
    '置換文字列。主に{}の中の文字を格納する
    Dim replaceValue As String
    '連想配列へのアクセスキー
    Dim key As String
    
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Charset = "UTF-8"
    st.Open
    
    '書き出し処理開始
    For i = 1 To objMetaSheet.Cells(Rows.Count, 1).End(xlUp).row
        writeText = objMetaSheet.Cells(i, 1).Value
            
        '波括弧がある場合は置換処理が必要（ここ正規表現チェックにしたい）
        If writeText Like "*{*" Then
            'Name項目をテキスト型にするときはdisplayFormatタグは不要
            If Not (writeText Like "*{表示形式}*" And objectInformation("データ型") = "Text") Then
            
                '例{表示ラベル（VBAは肯定先読みできないから前括弧は残る）
                replaceValue = regexpObj.Execute(writeText)(0)
                '例：表示ラベル
                key = Replace(replaceValue, "{", "")
                '例：<label>表示ラベル</label>
                writeText = Replace(writeText, "{", "")
                writeText = Replace(writeText, "}", "")
                '例：<label>hoge</label>
                writeText = Replace(writeText, key, objectInformation(key))
            End If
        End If
            
        st.writeText writeText & vbCrLf
    Next
    Call saveTextWithUTF8(st, fileName)

    
    'タブファイル作成
    If objectInformation("タブを作成") = True Then
    
        fileName = ThisWorkbook.Path & "\tabs\" & objApiName & ".tab-meta.xml"

        Set st = CreateObject("ADODB.Stream")
        st.Charset = "UTF-8"
        st.Open
        writeText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                    "<CustomTab xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf & _
                    "    <customObject>true</customObject>" & vbCrLf & _
                    "    <motif>Custom93: Shopping Cart</motif>" & vbCrLf & _
                    "</CustomTab>"

        st.writeText writeText & vbCrLf
        Call saveTextWithUTF8(st, fileName)
    End If
    
    'すべて表示のリストビュー作成
    fileName = ThisWorkbook.Path & "\objects\" & objApiName & "\listViews\All.listView-meta.xml"

    Set st = CreateObject("ADODB.Stream")
    st.Charset = "UTF-8"
    st.Open
    writeText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<ListView xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf & _
                "    <fullName>All</fullName>" & vbCrLf & _
                "    <filterScope>Everything</filterScope>" & vbCrLf & _
                "    <label>すべて選択</label>" & vbCrLf & _
                "</ListView>"

    st.writeText writeText & vbCrLf
    Call saveTextWithUTF8(st, fileName)
    
    MsgBox "完了しました。"
End Sub
'UTF-8で保存するときの保存処理をストリームobjectとファイル名で行う
Function saveTextWithUTF8(stream As Object, fileFullName As String)
        'Streamオブジェクトの先頭からの位置を指定する。Typeに値を設定するときは0である必要がある
        stream.Position = 0
        '扱うデータ種類をバイナリデータに変更する
        stream.Type = 1
        '読み取り開始位置？を3バイト目に移動する（3バイトはBOM付き部分を削除するため）
        stream.Position = 3
        'バイト文字を一時保存
        bytetmp = stream.Read
        'ここでは保存は不要。一度閉じて書き込んだ内容をリセットする目的がある
        stream.Close
        '再度開いて
        stream.Open
        'バイト形式で書き込むんで
        stream.write bytetmp
        '保存
        stream.SaveToFile fileFullName, 2
        'コピー先ファイルを閉じる
        stream.Close
End Function

