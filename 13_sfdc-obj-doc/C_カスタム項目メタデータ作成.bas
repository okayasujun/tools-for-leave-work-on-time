Attribute VB_Name = "C_カスタム項目メタデータ作成"
Dim itemSheet As Worksheet
Dim itemMetaSheet As Worksheet

Const PICK_LIST_PRE_TAG = "    " & "<valueSet>" & vbCrLf & _
                          "    " & "<restricted>true</restricted>" & vbCrLf & _
                          "    " & "<valueSetDefinition>" & vbCrLf & _
                          "    " & "    <sorted>false</sorted>" & vbCrLf
Const PICK_LIST_SUF_TAG = "    " & "        </valueSetDefinition>" & vbCrLf & _
                          "    " & "</valueSet>" & vbCrLf
'
'カスタム項目ごとにメタデータファイルを作成する
'データ型によって定義が必要なタグは[]シートに定義している
'
Sub C_カスタム項目メタデータ作成()

    Const PRE = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<CustomField xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf
    Const SUF = "</CustomField>"

    
    Set itemSheet = Sheets(ITEM_SHEET)
    Set itemMetaSheet = Sheets(ITEM_META_SHEET)
    
    Dim objApiName As String: objApiName = itemSheet.Cells(2, 4)
    Dim filePath As String: filePath = ThisWorkbook.Path & "\objects\" & objApiName & "\fields\"
    Dim fileName As String
    
    Dim lastRow As Integer: lastRow = itemSheet.Cells(4, 1).End(xlDown).row
    
    Dim labelName As String, itemName As String, apiName As String
    
    '###################################################
    '処理前にエラーチェックが必要
    '・API名の先頭が大文字があること
    '
    '###################################################
    
    
    For i = 5 To lastRow
        If itemSheet.Cells(i, 2).Value = "×" Then
            GoTo continue
        End If
        apiName = itemSheet.Cells(i, 5).Value
        
        fileName = filePath & apiName & ".field-meta.xml"
        
        'BOM削除
        With CreateObject("ADODB.Stream")
            .Charset = "UTF-8"
            .Open
            '書き出し処理開始
            .writeText PRE, 0
            .writeText getItemMetaData(i), 0
            .writeText SUF, 0
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
continue:
    Next
    
    MsgBox "完了しました。"
End Sub
Function getItemMetaData(row As Variant)
    Dim Value As String, returnValue As String, typeColumn As Integer, valueColumn As Integer
    Dim openTag As String, closeTag As String
    Dim dataType As String: dataType = itemSheet.Cells(row, 7).Value
    Dim valueType As String, listArray As Variant, listOneArray As Variant, listFlag As Boolean
    Dim writeTagFlag As Boolean: writeTagFlag = True
    
    '[CustomItem]シートより処理対象データタイプの列番号を取得する
    For i = 3 To 25
        If dataType = itemMetaSheet.Cells(2, i).Value Then
            typeColumn = i
            Exit For
        End If
    Next
    '縦方向ループ
    For i = 3 To 32
        valueColumn = itemMetaSheet.Cells(i, 2).Value
        If itemMetaSheet.Cells(i, typeColumn).Value = "〇" Then
            valueType = itemMetaSheet.Cells(i, 3).Value
            Value = itemSheet.Cells(row, valueColumn).Value
            
            If valueType = "テキスト" Then
            'TODO:デフォルト値の対応は必要
                'value = "&quot;" & value & "&quot;" '文字列のときはいるなあ・・・
            ElseIf valueType = "数値" Then
            ElseIf valueType = "真偽" Then
                Value = IIf(Value = "〇", "True", "False")
            ElseIf valueType = "リスト" Then
                Debug.Print Value
                listArray = Split(Value, vbLf)
                listFlag = True
                writeTagFlag = False
            End If
            
            If writeTagFlag Then
                openTag = itemMetaSheet.Cells(i, 1).Value
                closeTag = Replace(openTag, "<", "</")
                returnValue = returnValue & "    " & openTag & Value & closeTag & vbCrLf
            End If
            writeTagFlag = True
        ElseIf valueColumn = 15 Then
            '数式タグ
            If itemSheet.Cells(row, 8).Value = "〇" Then
                Value = itemSheet.Cells(row, 15).Value
                openTag = itemMetaSheet.Cells(i, 1).Value
                closeTag = Replace(openTag, "<", "</")
                returnValue = returnValue & "    " & openTag & Value & closeTag & vbCrLf
                
            End If
        ElseIf valueColumn = 16 Then
            '数式タグ
            If itemSheet.Cells(row, 8).Value = "〇" Then
                returnValue = returnValue & "    <formulaTreatBlanksAs>BlankAsZero</formulaTreatBlanksAs>" & vbCrLf
            End If
        End If
    Next
    '選択リストのタグ設定
    If listFlag Then
        returnValue = returnValue & PICK_LIST_PRE_TAG
        For Each Item In listArray
            listOneArray = Split(Item, ":")
            returnValue = returnValue & "    " & "<value>" & vbCrLf
            returnValue = returnValue & "    " & "<fullName>" & listOneArray(0) & "</fullName>" & vbCrLf
            returnValue = returnValue & "    " & "<default>false</default>" & vbCrLf
            returnValue = returnValue & "    " & "<label>" & listOneArray(1) & "</label>" & vbCrLf
            returnValue = returnValue & "    " & "</value>" & vbCrLf
        Next
        returnValue = returnValue & "    " & "        </valueSetDefinition>" & vbCrLf
        returnValue = returnValue & "    " & "</valueSet>" & vbCrLf
    End If
    getItemMetaData = returnValue
End Function
