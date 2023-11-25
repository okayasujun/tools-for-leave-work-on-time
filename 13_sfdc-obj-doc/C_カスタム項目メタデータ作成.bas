Attribute VB_Name = "C_カスタム項目メタデータ作成"
Const PICK_LIST_PRE_TAG = "    " & "<valueSet>" & vbCrLf & _
                          "    " & "<restricted>true</restricted>" & vbCrLf & _
                          "    " & "<valueSetDefinition>" & vbCrLf & _
                          "    " & "    <sorted>false</sorted>" & vbCrLf
Const PICK_LIST_SUF_TAG = "    " & "        </valueSetDefinition>" & vbCrLf & _
                          "    " & "</valueSet>" & vbCrLf
Const PRE = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
            "<CustomField xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf
Const SUF = "</CustomField>"
'
'カスタム項目ごとにメタデータファイルを作成する
'データ型によって定義が必要なタグは[]シートに定義している
'
Sub C_カスタム項目メタデータ作成()
    Call initiarize

    '最終行
    Dim lastRow As Integer: lastRow = itemSheet.Cells(4, 1).End(xlDown).row
    'ラベル名、API名
    Dim labelName As String
    Dim apiName As String
    
    '###################################################
    '処理前にエラーチェックが必要
    '・API名の先頭が大文字があること
    '
    '###################################################
    
    Dim i As Integer
    For i = 5 To lastRow
        '有効項目のみ対象とする
        If itemSheet.Cells(i, 2).Value = "〇" Then
 
            apiName = itemSheet.Cells(i, 5).Value
            fileName = fieldsDirPath & apiName & ".field-meta.xml"

            Call openStream
            '書き出し処理開始
            stream.writeText PRE, 0
            stream.writeText getItemMetaData(i), 0
            stream.writeText SUF, 0
            Call saveTextWithUTF8(stream, fileName)
        End If
    Next
    
    MsgBox "完了しました。"
End Sub
'メタデータのテキスト情報を返す
Function getItemMetaData(row As Integer)
    '出力テキストのタグ以外の部分
    Dim writeText As String
    'モドリッチ
    Dim returnValue As String
    '[CustomItem]シートにおいて、対象のデータ型が記載してある列番号
    Dim dataTypeColumn As Integer
    '[項目]シート上の設定列を保持する番号
    Dim valueColumn As Integer
    '[CustomItem]シートの1列目の値を保持する
    Dim openTag As String
    '閉じるタグをopenTagから作って格納する
    Dim closeTag As String
    'データ型（日本語）
    Dim dataType As String: dataType = itemSheet.Cells(row, 7).Value
    '各属性の設定値のデータ型
    Dim valueType As String
    '選択リスト値リスト
    Dim listArray As Variant
    '選択肢ひとつのAPI名、ラベルの配列
    Dim listOneArray As Variant
    '選択リスト選択肢のメタ情報を書き出すぜフラグ
    Dim listFlag As Boolean
    'メタデータファイルに書き出すかどうかを「〇」の有無で取得する
    Dim writeTagFlag As Boolean: writeTagFlag = True
    
    dataType = itemSheet.Cells(row, 7).Value
    dataType = IIf(itemSheet.Cells(row, 8).Value = "〇", "(数式)" & dataType, dataType)
    
    '走査中行のデータ型を[CustomItem]シートから探す
    For i = 4 To 31
        If dataType = itemMetaSheet.Cells(2, i).Value Then
            dataTypeColumn = i
            Exit For
        End If
    Next
    
    'メタ情報の生成
    For i = 3 To 37
        valueColumn = itemMetaSheet.Cells(i, 2).Value
        
        If itemMetaSheet.Cells(i, dataTypeColumn).Value And valueColumn > 0 Then
            valueType = itemMetaSheet.Cells(i, 3).Value
            writeText = itemSheet.Cells(row, valueColumn).Value
            
            '項目が確定している場合の補助処理
            If valueType = "テキスト" Then
            
'                If itemMetaSheet.Cells(i, 1).Value = "<defaultValue>" And writeText <> "" Then
'                    writeText = "&quot;" & writeText & "&quot;" '文字列のときはいるなあ・・・
'                End If
'                TODO:デフォルト値の対応は必要（やり方の検討からして）

            ElseIf valueType = "数値" Then
                
            ElseIf valueType = "真偽" Then
                writeText = IIf(writeText = "〇", "True", "False")
                
            ElseIf valueType = "リスト" Then
                'Debug.Print writeText
                listArray = Split(writeText, vbLf)
                listFlag = True
                writeTagFlag = False
                
            ElseIf valueType = "BlankAsZero" Then
                '「BlankAsZero」以外の場合
                writeTagFlag = writeText = "BlankAsZero"
                
            End If
            
            If writeTagFlag Then
                openTag = itemMetaSheet.Cells(i, 1).Value
                closeTag = Replace(openTag, "<", "</")
                returnValue = returnValue & "    " & openTag & writeText & closeTag & vbCrLf
            End If
            writeTagFlag = True
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

