Attribute VB_Name = "B_オブジェクトメタデータ作成"
Sub A_オブジェクトメタデータ作成()
    Call initiarize
    fileName = ThisWorkbook.path & "\objects\" & objApiName & "\" & objApiName & ".object-meta.xml"
    
    '連想配列作成（A,D列を対応付ける）
    Dim objectInformation As Object
    Set objectInformation = CreateObject("Scripting.Dictionary")
    For i = 1 To objSheet.Cells(Rows.Count, 1).End(xlUp).row
        objectInformation(objSheet.Cells(i, 1).Value) = objSheet.Cells(i, 4).Value
    Next
    
    '正規表現（置換のため）
    Call setupRegexp("{.*(?=})")
    
    'ファイルに書き出すテキスト
    Dim writeText As String
    '置換文字列。主に{}の中の文字を格納する
    Dim replaceValue As String
    '連想配列へのアクセスキー
    Dim key As String
    
    'テキストファイル出力準備
    Call openStream
    
    '書き出し処理開始
    For i = 1 To objMetaSheet.Cells(Rows.Count, 1).End(xlUp).row
        writeText = objMetaSheet.Cells(i, 1).Value
            
        '波括弧がある場合は置換処理が必要（ここ正規表現チェックにしたい）
        If writeText Like "*{*" Then
            'Name項目をテキスト型にするときはdisplayFormatタグは不要
            If Not (writeText Like "*{表示形式}*" And objectInformation("データ型") = "Text") Then
            
                '例{表示ラベル（VBAは肯定先読みできないから前括弧は残る）
                replaceValue = regexp.Execute(writeText)(0)
                '例：表示ラベル
                key = Replace(replaceValue, "{", "")
                '例：<label>表示ラベル</label> ここなんともだせーな
                writeText = Replace(writeText, "{", "")
                writeText = Replace(writeText, "}", "")
                '例：<label>hoge</label>
                writeText = Replace(writeText, key, objectInformation(key))
            End If
        End If
            
        stream.writeText writeText & vbCrLf
    Next
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "完了しました。"
End Sub

Sub B_タブメタデータ作成()
    Call initiarize
    fileName = ThisWorkbook.path & "\tabs\" & objApiName & ".tab-meta.xml"

    'テキストファイル出力準備
    Call openStream
    'ファイルに書き出すテキスト
    Dim writeText As String: writeText = _
        "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<CustomTab xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf & _
        "    <customObject>true</customObject>" & vbCrLf & _
        "    <motif>Custom93: Shopping Cart</motif>" & vbCrLf & _
        "</CustomTab>"

    stream.writeText writeText & vbCrLf
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "完了しました。"
End Sub

Sub C_リストビューメタデータ作成()
    Call initiarize
    fileName = ThisWorkbook.path & "\objects\" & objApiName & "\listViews\All.listView-meta.xml"

    Call openStream
    'ファイルに書き出すテキスト
    Dim writeText As String: writeText = _
        "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<ListView xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf & _
        "    <fullName>All</fullName>" & vbCrLf & _
        "    <filterScope>Everything</filterScope>" & vbCrLf & _
        "    <label>すべて選択</label>" & vbCrLf & _
        "</ListView>"

    stream.writeText writeText & vbCrLf
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "完了しました。"
End Sub

