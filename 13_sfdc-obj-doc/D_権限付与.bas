Attribute VB_Name = "D_�����t�^"
Sub D_�����t�^()
    Call initiarize
    'TODO:�����͂܂������Ƃ��
    fileName = ThisWorkbook.path & "\objects\" & objApiName & "\" & objApiName & ".csv"

    '�e�L�X�g�t�@�C���o�͏���
    Call openStream
    '�w�b�_���ݒ�
    stream.writeText "PARENTID,SOBJECTTYPE,FIELD,PERMISSIONSREAD,PERMISSIONSEDIT" & vbCrLf
    '�����o�������J�n
    For i = 8 To permissionSheet.Cells(13, Columns.Count).End(xlToLeft).column Step 2
        For j = 14 To permissionSheet.Cells(Rows.Count, 1).End(xlUp).row
            writeText = permissionSheet.Cells(4, i).Value & ","
            writeText = writeText & permissionSheet.Cells(3, 2).Value & ","
            writeText = writeText & permissionSheet.Cells(j, 5).Value & ","
            writeText = writeText & permissionSheet.Cells(j, i).Value & ","
            writeText = writeText & permissionSheet.Cells(j, i + 1).Value
            'TODO:�w�b�_�s�̃X�L�b�v
            'TODO:�K�{���ڂ̍l��
            'TODO:Name���ڂ̍l��
            'TODO:�������ڂ̍l��
            ''TODO:�uFIELD�v���ڂ̓I�u�W�F�N�g��API�{����API��
            stream.writeText writeText & vbCrLf
        Next
    Next
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "�������܂����B"
End Sub
