Attribute VB_Name = "D_�����t�^"
Sub D_���ڌ����t�^()
    Call initiarize
    fileName = ThisWorkbook.path & "\permission\" & objApiName & "\���ڌ���.csv"
    objApiName = permissionSheet.Cells(3, 2).Value

    '�e�L�X�g�t�@�C���o�͏���
    Call openStream
    '�w�b�_���ݒ�
    stream.writeText "PARENTID,SOBJECTTYPE,FIELD,PERMISSIONSREAD,PERMISSIONSEDIT" & vbCrLf
    '�����o�������J�n
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

                'TODO:�������ڂ�������Q�Ƃ̂݉\
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
            'TODO:i���P���ڂ̎������A�Q�O�s���ƂɃw�b�_�s��}�����鏈�������Ă��E�E�E
        Next
    Next
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "�������܂����B"
End Sub
'�����o�͉\�ȍ��ڂ��ǂ�����Ԃ�
Function isWritableItemPermission(index As Integer)
    isWritableItemPermission = True
    
    '�w�b�_�s����Ȃ����`�F�b�N
    If permissionSheet.Cells(index, 1).Interior.Color = NO_COLOR Then
    
        '�L�����`�F�b�N
        If permissionSheet.Cells(index, 2).Value = "�~" Then
            isWritableItemPermission = isWritableItemPermission And False
        End If
        'Name���ڂ��`�F�b�N
        If permissionSheet.Cells(index, 5).Value = "Name" Then
            isWritableItemPermission = isWritableItemPermission And False
        End If
        '��]���ڂ��`�F�b�N
        If permissionSheet.Cells(index, 6).Value = "��]�֌W" Then
            isWritableItemPermission = isWritableItemPermission And False
        End If
        '�K�{���ڂ��`�F�b�N
        If permissionSheet.Cells(index, 8).Value Then
            isWritableItemPermission = isWritableItemPermission And False
        End If
    Else
        isWritableItemPermission = isWritableItemPermission And False
    End If
End Function
'�������ڂ��`�F�b�N����
Function isFormulaItem(index As Integer)
    isFormulaItem = False
    If permissionSheet.Cells(index, 1).Interior.Color = NO_COLOR Then
        If permissionSheet.Cells(index, 7).Value Then
            isFormulaItem = True
        End If
    End If
End Function
Sub E_�I�u�W�F�N�g�����t�^()
    Call initiarize
    fileName = ThisWorkbook.path & "\permission\" & objApiName & "\�I�u�W�F�N�g����.csv"
    objApiName = permissionSheet.Cells(3, 2).Value
    Dim writeText As String
    '�e�L�X�g�t�@�C���o�͏���
    Call openStream
    '�w�b�_���ݒ�
    stream.writeText "PARENTID,SOBJECTTYPE,PERMISSIONSREAD,PERMISSIONSCREATE,PERMISSIONSEDIT,PERMISSIONSDELETE,PERMISSIONSVIEWALLRECORDS,PERMISSIONSMODIFYALLRECORDS" & vbCrLf
    '�����o�������J�n
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
    MsgBox "�������܂����B"
End Sub
Sub F_�^�u�ݒ�()
    Call initiarize
    fileName = ThisWorkbook.path & "\permission\" & objApiName & "\�^�u�ݒ�.csv"
    objApiName = permissionSheet.Cells(3, 2).Value
    Dim writeText As String
    '�e�L�X�g�t�@�C���o�͏���
    Call openStream
    '�w�b�_���ݒ�
    stream.writeText "NAME,PARENTID,VISIBILITY" & vbCrLf
    '�����o�������J�n
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
    MsgBox "�������܂����B"
End Sub
