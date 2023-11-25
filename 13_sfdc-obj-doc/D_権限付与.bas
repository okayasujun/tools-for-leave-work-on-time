Attribute VB_Name = "D_権限付与"
Sub D_項目権限付与()
    Call initiarize
    fileName = ThisWorkbook.path & "\permission\" & objApiName & "\項目権限.csv"
    objApiName = permissionSheet.Cells(3, 2).Value

    'テキストファイル出力準備
    Call openStream
    'ヘッダ情報設定
    stream.writeText "PARENTID,SOBJECTTYPE,FIELD,PERMISSIONSREAD,PERMISSIONSEDIT" & vbCrLf
    '書き出し処理開始
    For i = 9 To permissionSheet.Cells(13, Columns.Count).End(xlToLeft).column Step 2
        For j = 14 To permissionSheet.Cells(Rows.Count, 1).End(xlUp).row
            If isWritableItemPermission(CInt(j)) Then
                'ParentId
                writeText = permissionSheet.Cells(4, i).Value & ","
                'SobjectType
                writeText = writeText & objApiName & ","
                'Field
                writeText = writeText & objApiName & "." & permissionSheet.Cells(j, 5).Value & ","
                'PermissionsRead
                writeText = writeText & permissionSheet.Cells(j, i).Value & ","
                'PermissionsEdit
                writeText = writeText & permissionSheet.Cells(j, i + 1).Value

                'TODO:数式項目だったら参照のみ可能
                stream.writeText writeText & vbCrLf
            ElseIf permissionSheet.Cells(j, 1).Interior.Color = NO_COLOR Then
                permissionSheet.Cells(j, i).Font.Color = RGB(191, 191, 191)
                permissionSheet.Cells(j, i + 1).Font.Color = RGB(191, 191, 191)
            End If
            
            If isFormulaItem(CInt(j)) Then
                permissionSheet.Cells(j, i + 1).Font.Color = RGB(191, 191, 191)
            Else
                permissionSheet.Cells(j, i + 1).Font.Color = RGB(0, 0, 0)
            End If
            'TODO:iが１周目の時だけ、２０行ごとにヘッダ行を挿入する処理書きてえ・・・
        Next
    Next
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "完了しました。"
End Sub
'権限出力可能な項目かどうかを返す
Function isWritableItemPermission(index As Integer)
    isWritableItemPermission = True
    
    'ヘッダ行じゃないかチェック
    If permissionSheet.Cells(index, 1).Interior.Color = NO_COLOR Then
    
        '有効かチェック
        If permissionSheet.Cells(index, 2).Value = "×" Then
            isWritableItemPermission = isWritableItemPermission And False
        End If
        'Name項目かチェック
        If permissionSheet.Cells(index, 5).Value = "Name" Then
            isWritableItemPermission = isWritableItemPermission And False
        End If
        '主従項目かチェック
        If permissionSheet.Cells(index, 6).Value = "主従関係" Then
            isWritableItemPermission = isWritableItemPermission And False
        End If
        '必須項目かチェック
        If permissionSheet.Cells(index, 8).Value Then
            isWritableItemPermission = isWritableItemPermission And False
        End If
    Else
        isWritableItemPermission = isWritableItemPermission And False
    End If
End Function
'数式項目かチェックする
Function isFormulaItem(index As Integer)
    isFormulaItem = False
    If permissionSheet.Cells(index, 1).Interior.Color = NO_COLOR Then
        If permissionSheet.Cells(index, 7).Value Then
            isFormulaItem = True
        End If
    End If
End Function
Sub E_オブジェクト権限付与()
    Call initiarize
    fileName = ThisWorkbook.path & "\permission\" & objApiName & "\オブジェクト権限.csv"
    objApiName = permissionSheet.Cells(3, 2).Value
    Dim writeText As String
    'テキストファイル出力準備
    Call openStream
    'ヘッダ情報設定
    stream.writeText "PARENTID,SOBJECTTYPE,PERMISSIONSREAD,PERMISSIONSCREATE,PERMISSIONSEDIT,PERMISSIONSDELETE,PERMISSIONSVIEWALLRECORDS,PERMISSIONSMODIFYALLRECORDS" & vbCrLf
    '書き出し処理開始
    For i = 9 To permissionSheet.Cells(2, Columns.Count).End(xlToLeft).column Step 2
        'ParentId
        writeText = permissionSheet.Cells(4, i).Value & ","
        'SobjectType
        writeText = writeText & objApiName & ","
        For j = 6 To 11
            writeText = writeText & permissionSheet.Cells(j, i).Value & ","
        Next
        stream.writeText deleteEndText(writeText) & vbCrLf
    Next
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "完了しました。"
End Sub
Sub F_タブ設定()
    Call initiarize
    fileName = ThisWorkbook.path & "\permission\" & objApiName & "\タブ設定.csv"
    objApiName = permissionSheet.Cells(3, 2).Value
    Dim writeText As String
    'テキストファイル出力準備
    Call openStream
    'ヘッダ情報設定
    stream.writeText "NAME,PARENTID,VISIBILITY" & vbCrLf
    '書き出し処理開始
    For i = 9 To permissionSheet.Cells(2, Columns.Count).End(xlToLeft).column Step 2
        'SobjectType
        writeText = objApiName & ","
        'ParentId
        writeText = writeText & permissionSheet.Cells(4, i).Value & ","
        'Visibility
        writeText = writeText & permissionSheet.Cells(5, i + 1).Value & ","
        
        
        stream.writeText deleteEndText(writeText) & vbCrLf
    Next
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "完了しました。"
End Sub
