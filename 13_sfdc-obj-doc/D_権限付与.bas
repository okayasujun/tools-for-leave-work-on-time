Attribute VB_Name = "D_権限付与"
Sub D_権限付与()
    Call initiarize
    'TODO:ここはまたちゃんとやる
    fileName = ThisWorkbook.path & "\objects\" & objApiName & "\" & objApiName & ".csv"

    'テキストファイル出力準備
    Call openStream
    'ヘッダ情報設定
    stream.writeText "PARENTID,SOBJECTTYPE,FIELD,PERMISSIONSREAD,PERMISSIONSEDIT" & vbCrLf
    '書き出し処理開始
    For i = 8 To permissionSheet.Cells(13, Columns.Count).End(xlToLeft).column Step 2
        For j = 14 To permissionSheet.Cells(Rows.Count, 1).End(xlUp).row
            writeText = permissionSheet.Cells(4, i).Value & ","
            writeText = writeText & permissionSheet.Cells(3, 2).Value & ","
            writeText = writeText & permissionSheet.Cells(j, 5).Value & ","
            writeText = writeText & permissionSheet.Cells(j, i).Value & ","
            writeText = writeText & permissionSheet.Cells(j, i + 1).Value
            'TODO:ヘッダ行のスキップ
            'TODO:必須項目の考慮
            'TODO:Name項目の考慮
            'TODO:数式項目の考慮
            ''TODO:「FIELD」項目はオブジェクトのAPI＋項目APIだ
            stream.writeText writeText & vbCrLf
        Next
    Next
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "完了しました。"
End Sub
