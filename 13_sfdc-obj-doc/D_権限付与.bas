Attribute VB_Name = "D_権限付与"
Dim itemSheet As Worksheet
Sub D_権限付与()

    Const PRE = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<Profile xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf
    Const SUF = "</Profile>"

'自動採番は編集不可
'必須項目は参照・編集が〇であること
'など、最初にチェック処理が必要
    
    Set itemSheet = Sheets(ITEM_SHEET)
    
    Dim objApiName As String: objApiName = itemSheet.Cells(2, 4)
    Dim filePath As String: filePath = ThisWorkbook.Path & "\profiles\Admin.profile-meta.xml"
    Dim fileName As String: fileName = filePath '本来はここで、ファイル名設定する
    
    Dim lastRow As Integer: lastRow = itemSheet.Cells(4, 1).End(xlDown).row
    
    Dim itemApiName As String, readPermission As String, editPermission As String
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        
        .writeText PRE
        For i = 5 To lastRow
            If itemSheet.Cells(i, 17) = "〇" Then
                GoTo continue
            End If
            itemApiName = itemSheet.Cells(i, 5).Value
            editPermission = itemSheet.Cells(i, 40) = "〇"
            readPermission = itemSheet.Cells(i, 39) = "〇"
            '38,39
            .writeText "    <fieldPermissions>" & vbCrLf
            .writeText "        <editable>" & editPermission & "</editable>" & vbCrLf
            .writeText "        <field>" & objApiName & "." & itemApiName & "</field>" & vbCrLf
            .writeText "        <readable>" & readPermission & "</readable>" & vbCrLf
            .writeText "    </fieldPermissions>" & vbCrLf
continue:
        Next
        'タブの設定
        .writeText "    <tabVisibilities>" & vbCrLf
        .writeText "        <tab>" & objApiName & "</tab>" & vbCrLf
        .writeText "        <visibility>" & itemSheet.Cells(3, 40) & "</visibility>" & vbCrLf
        .writeText "    </tabVisibilities>" & vbCrLf
        
        .writeText SUF
        '書き出し処理終了
        .Position = 0
        .Type = 1
        .Position = 3
        bytetmp = .Read
        .SaveToFile fileName, 2
        'コピー先ファイルを閉じる
        .Close
    End With
    
    'UTF-8でテキストファイルへ出力する
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .LineSeparator = 10
        .Type = 1
        .Open
        .write bytetmp
        .SetEOS
        .SaveToFile fileName, 2
        .Close
    End With
    
    MsgBox "完了しました。"
End Sub
