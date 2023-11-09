Attribute VB_Name = "X_ページレイアウト作成"
Dim layoutSheet As Worksheet
Sub E_ページレイアウト作成()

    Const PRE = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<Layout xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf & _
                "    <excludeButtons>Submit</excludeButtons>" & vbCrLf
    Const SUF = "</Layout>"

    
    Set layoutSheet = Sheets(LAYOUT_SHEET)
    
    Dim objApiName As String: objApiName = layoutSheet.Cells(2, 3)
    Dim filePath As String: filePath = ThisWorkbook.path & "\layouts\" & objApiName & "-レイアウト名" & ".layout-meta.xml"
    Dim fileName As String: fileName = filePath '本来はここで、ファイル名設定する
    
    Dim lastRow As Integer: lastRow = layoutSheet.Cells(layoutSheet.Rows.Count, 1).End(xlUp).row
    
    Dim itemApiName As String, readPermission As String, editPermission As String
    
    Dim layoutMap As Object
    Set layoutMap = CreateObject("Scripting.Dictionary")
    Dim sectionMap As Object
    Set sectionMap = CreateObject("Scripting.Dictionary")
    
    Dim layoutCount As Integer
    'セクションセルの色
    Dim sectionColor As Long: sectionColor = layoutSheet.Cells(5, 2).Interior.Color
    
    For i = 4 To lastRow
        If layoutSheet.Cells(i, 1) <> "" Then
            layoutMap(i) = layoutSheet.Cells(i, 1).Value
            'layoutMap(layoutCount) = layoutSheet.Cells(i, 1).value
            'layoutCount = layoutCount + 1
        End If
        'まずレイアウト名をマップにとる
        '次にセクションをマップに取る
        'セクション内で項目ごとに構築していく
    
    Next
    
    'この辺りはロジック再考です。レイアウトとセクションを別のループで取得する必要があるのかどうか。
    For Each l In layoutMap
        For i = l To l + 100
            If layoutSheet.Cells(i, 2).Interior.Color = sectionColor Then
                sectionMap(i) = layoutSheet.Cells(i, 2).Value
            End If
        Next
    Next
    Dim roopToValue As Integer
    Dim sectionKeys
    sectionKeys = sectionMap.keys
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .writeText PRE
        
        'ここに本当はもう一段階ループがあって複数のレイアウトに対応する予定。
        sectionCount = 0
        For Each l In sectionMap
            .writeText "    <layoutSections>" & vbCrLf
            .writeText "        <customLabel>true</customLabel>" & vbCrLf
            .writeText "        <detailHeading>" & layoutSheet.Cells(l, 5) & "</detailHeading>" & vbCrLf
            .writeText "        <editHeading>" & layoutSheet.Cells(l, 9) & "</editHeading>" & vbCrLf
            .writeText "        <label>" & layoutSheet.Cells(l, 2) & "</label>" & vbCrLf
            .writeText "        <layoutColumns>" & vbCrLf
            If UBound(sectionKeys) = sectionCount Then
                roopToValue = 22
            Else
                roopToValue = sectionKeys(sectionCount + 1) - 1
            End If
            For i = l + 1 To roopToValue
                If layoutSheet.Cells(i, 3) = "空白" Then
                    .writeText "            <layoutItems>" & vbCrLf
                    .writeText "                <emptySpace>true</emptySpace>" & vbCrLf
                    .writeText "            </layoutItems>" & vbCrLf
                ElseIf layoutSheet.Cells(i, 3) <> "" Then
                    .writeText "            <layoutItems>" & vbCrLf
                    .writeText "                <behavior>" & layoutSheet.Cells(i, 5) & "</behavior>" & vbCrLf
                    .writeText "                <field>" & layoutSheet.Cells(i, 3) & "</field>" & vbCrLf
                    .writeText "            </layoutItems>" & vbCrLf
                End If
            Next
            .writeText "        </layoutColumns>" & vbCrLf
            .writeText "        <layoutColumns>" & vbCrLf
            For i = l + 1 To roopToValue
                If layoutSheet.Cells(i, 7) = "空白" Then
                    .writeText "            <layoutItems>" & vbCrLf
                    .writeText "                <emptySpace>true</emptySpace>" & vbCrLf
                    .writeText "            </layoutItems>" & vbCrLf
                ElseIf layoutSheet.Cells(i, 7) <> "" Then
                    .writeText "            <layoutItems>" & vbCrLf
                    .writeText "                <behavior>" & layoutSheet.Cells(i, 9) & "</behavior>" & vbCrLf
                    .writeText "                <field>" & layoutSheet.Cells(i, 7) & "</field>" & vbCrLf
                    .writeText "            </layoutItems>" & vbCrLf
                End If
            Next
            .writeText "        </layoutColumns>" & vbCrLf
            .writeText "        <style>TwoColumnsTopToBottom</style>" & vbCrLf
            .writeText "    </layoutSections>" & vbCrLf
            sectionCount = sectionCount + 1
        Next
        
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
