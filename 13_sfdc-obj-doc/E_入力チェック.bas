Attribute VB_Name = "E_���̓`�F�b�N"
Private errorSheet As Worksheet
Private logSheet As Worksheet
Private lastRow As Integer
Private lastColumn As Integer
Private dataType As String
Private errorCheckColumn As Integer
Private errorMsgFormat As String
Private errorLastRow As Integer
Private condition As String
Private logWriteLine As Integer
Private i As Integer
Private j As Integer
Sub G_���̓`�F�b�N()
    '��TODO
    '�EAPI���̃��[���s���s���Ȃ���
    '�EAPI���ɏd���͂Ȃ���
    '�E�I�u�W�F�N�g�̐ݒ�Ɛ������Ă��邩
    '�EAPI���̓������͑啶���ł���i����͂����ȁj
    '�E�I�u�W�F�N�g�̋��L�ݒ�Ǝ�]�֌W���ڂ̗L��
    Call initiarize
    Call init
    Call repaint
    
    For i = 5 To lastRow
        With itemSheet
            dataType = .Cells(i, 7).Value
            dataType = IIf(.Cells(i, 8).Value = ON_TRUE, "(����)" & dataType, dataType)
            
            For j = 2 To errorSheet.Cells(1, 1).End(xlDown).row
                condition = errorSheet.Cells(j, 6).Value
                errorCheckColumn = errorSheet.Cells(j, 4).Value
                errorMsgFormat = errorSheet.Cells(j, 7).Value
            
                '�L�����ڂŁA�f�[�^�^����v����΃`�F�b�N�ɂ�����
                If .Cells(i, 2).Value = ON_TRUE And errorSheet.Cells(j, 2).Value = dataType Then
                
                    If condition = "�ȏ�" Then
                        If Not .Cells(i, errorCheckColumn).Value >= errorSheet.Cells(j, 5).Value Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    ElseIf condition = "�ȉ�" Then
                        If Not .Cells(i, errorCheckColumn).Value <= errorSheet.Cells(j, 5).Value Then
                            Call setError(errorMsgFormat)
                        End If
                    
                    ElseIf condition = "������" Then
                        If Not .Cells(i, errorCheckColumn).Value = errorSheet.Cells(j, 5).Value Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    ElseIf condition = "�K�{" Then
                        If Not .Cells(i, errorCheckColumn).Value <> "" Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    ElseIf condition = "���K�\���Ɉ�v����" And dataType = "�I�����X�g" Then
                        Call setupRegexp(errorSheet.Cells(j, 5))
                        If Not UBound(Split(.Cells(i, errorCheckColumn).Value, vbLf)) + 1 = regexp.Execute(.Cells(i, errorCheckColumn).Value).Count Then
                            Call setError(errorMsgFormat)
                            logSheet.Cells(logWriteLine - 1, 2).WrapText = False
                        End If
                        
                    ElseIf condition = "�ƈ�v����" Then
                        Dim targetArray() As String: targetArray = Split(errorSheet.Cells(j, 5).Value, ",")
                        If Not isExistArray(targetArray, .Cells(i, errorCheckColumn).Value) Then
                            Call setError(errorMsgFormat)
                        End If
                    End If
                    
                '[Error��`]��A�f�[�^�^���w��̃P�[�X
                ElseIf .Cells(i, 2).Value = "�Z" And errorSheet.Cells(j, 2).Value = "-" Then
                    If errorSheet.Cells(j, 5) = "���s����" And condition = "�܂܂Ȃ�" Then
                        If .Cells(i, errorCheckColumn).Value Like "*" & vbLf & "*" _
                            Or .Cells(i, errorCheckColumn).Value Like "*" & vbCrLf & "*" _
                            Or .Cells(i, errorCheckColumn).Value Like "*" & vbCr & "*" Then
                            Call setError(errorMsgFormat)
                            logSheet.Cells(logWriteLine - 1, 2).WrapText = False
                        End If
                        
                    ElseIf condition = "�K�{" Then
                        If Not .Cells(i, errorCheckColumn).Value <> "" Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    ElseIf condition = "���ȉ�" Then
                        If Not Len(.Cells(i, errorCheckColumn).Value) <= errorSheet.Cells(j, 5).Value Then
                            Call setError(errorMsgFormat)
                        End If
                    
                    ElseIf condition = "�ƈ�v����" Then
                        targetArray = Split(errorSheet.Cells(j, 5).Value, ",")
                        If Not isExistArray(targetArray, .Cells(i, errorCheckColumn).Value) Then
                            Call setError(errorMsgFormat)
                        End If
                        
                    End If
                End If
            Next
        End With
    Next
    
    If logWriteLine = 2 Then
        MsgBox "���̓`�F�b�N���������܂����B" & vbCrLf & "���͕s���͂���܂���B"
    Else
        MsgBox "���͕s��������܂��B[log]�V�[�g���m�F���Ă��������B"
        logSheet.Select
    End If
End Sub
'�G���[����ݒ肷��
Private Function setError(format As String)
    logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
    logSheet.Cells(logWriteLine, 2) = getMessage(format, _
                                i, itemSheet.Cells(i, 3), errorSheet.Cells(j, 3), _
                                errorSheet.Cells(j, 5), errorSheet.Cells(j, 6), _
                                itemSheet.Cells(i, errorCheckColumn).Value)
    logWriteLine = logWriteLine + 1
    '�ΏۃZ���n�C���C�g
    itemSheet.Cells(i, errorCheckColumn).Interior.Color = RGB(255, 255, 0)
End Function
'�`�F�b�N�����p����������
Private Function init()
    Set errorSheet = Sheets("Error��`")
    Set logSheet = Sheets("log")
    lastRow = itemSheet.Cells(5, 1).End(xlDown).row
    lastColumn = itemSheet.Cells(4, 1).End(xlToRight).column
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 1).Font.Bold = True
    logSheet.Cells(1, 2) = "���e"
    logSheet.Cells(1, 2).Font.Bold = True
    logWriteLine = 2
End Function
'�Z���h�F��K�؂ȏ�Ԃɂ���
Function repaint()

    For i = 5 To lastRow
        With itemSheet
            dataType = .Cells(i, 7).Value
            dataType = IIf(.Cells(i, 8).Value = "�Z", "(����)" & dataType, dataType)
            
            If .Cells(i, 2).Value = "�~" Then
                '���ׂăO���[�ɓh��
                .Range(.Cells(i, 1), .Cells(i, 38)).Interior.Color = RGB(191, 191, 191)
            Else
                '�F����
                '�������s�̃f�[�^�^���`�V�[�g����T��
                For j = 4 To 31
                    If dataType = itemMetaSheet.Cells(2, j).Value Then
                        errorCheckColumn = j
                        Exit For
                    End If
                Next
                    
                '�f�[�^�^�ɉ����ċL�ڕs�v�ȏ����O���[�ɓh�F����
                For j = 3 To 37
                    colorChangeColumn = itemMetaSheet.Cells(j, 2).Value
                    If itemMetaSheet.Cells(j, errorCheckColumn).Value And colorChangeColumn > 0 Then
                        '�L���Z��
                        .Cells(i, colorChangeColumn).Interior.Color = NO_COLOR
                    ElseIf colorChangeColumn > 0 Then
                        '�����Z��
                        .Cells(i, colorChangeColumn).Interior.Color = RGB(191, 191, 191)
                        .Cells(i, colorChangeColumn).Value = ""
                    End If
                Next
                '��L�����ŋd���ĂȂ���
                .Cells(i, 1).Interior.Color = NO_COLOR
                .Cells(i, 2).Interior.Color = NO_COLOR
                .Cells(i, 4).Interior.Color = NO_COLOR
                .Cells(i, 6).Interior.Color = NO_COLOR
                .Cells(i, 7).Interior.Color = NO_COLOR
                .Cells(i, 8).Interior.Color = NO_COLOR
            End If
        End With
    Next

End Function

