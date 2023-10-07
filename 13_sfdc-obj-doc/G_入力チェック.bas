Attribute VB_Name = "G_���̓`�F�b�N"
Private itemSheet As Worksheet
Private itemMetaSheet As Worksheet
Private errorSheet As Worksheet
Private logSheet As Worksheet
Private lastRow As Integer
Private lastColumn As Integer
Private dataType As String
Private dataTypeColumn As Integer
Private errorLastRow As Integer
Private condition As String
Private logWriteLine As Integer
Const NO_COLOR = 16777215
Sub G_���̓`�F�b�N()
    '�EAPI���̃��[���s���s���Ȃ���
    '�EAPI���ɏd���͂Ȃ���
    '�E�I�u�W�F�N�g�̐ݒ�Ɛ������Ă��邩
    '�EAPI���̓������͑啶���ł���i����͂����ȁj
    '�E�I�u�W�F�N�g�̋��L�ݒ�Ǝ�]�֌W���ڂ̗L��
    Call init
    Call repaint
    
    For i = 5 To lastRow
        With itemSheet
            dataType = .Cells(i, 7).Value
            dataType = IIf(.Cells(i, 8).Value = "�Z", "(����)" & dataType, dataType)
            
            For j = 2 To errorSheet.Cells(1, 1).End(xlDown).row
                    condition = errorSheet.Cells(j, 6).Value
                    dataTypeColumn = errorSheet.Cells(j, 4).Value
            
                If .Cells(i, 2).Value = "�Z" And errorSheet.Cells(j, 2).Value = dataType Then
                
                    If condition = "�ȏ�" Then
                        If Not .Cells(i, dataTypeColumn).Value >= errorSheet.Cells(j, 5).Value Then
                            '����G���[
                            logSheet.Cells(logWriteLine, 1) = i & "�s�ڂ́u" & .Cells(i, 3) & "�v���ڂ́u" _
                                    & errorSheet.Cells(j, 3) & "�v��" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "�ɂ��Ă��������B(" & .Cells(i, dataTypeColumn).Value & ")"
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
                        End If
                        
                    ElseIf condition = "�ȉ�" Then
                        If Not .Cells(i, dataTypeColumn).Value <= errorSheet.Cells(j, 5).Value Then
                            '����G���[
                            logSheet.Cells(logWriteLine, 1) = i & "�s�ڂ́u" & .Cells(i, 3) & "�v���ڂ́u" _
                                    & errorSheet.Cells(j, 3) & "�v��" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "�ɂ��Ă��������B(" & .Cells(i, dataTypeColumn).Value & ")"
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
                        End If
                    
                    ElseIf condition = "������" Then
                        If Not .Cells(i, dataTypeColumn).Value = errorSheet.Cells(j, 5).Value Then
                            logSheet.Cells(logWriteLine, 1) = i & "�s�ڂ́u" & .Cells(i, 3) & "�v���ڂ́u" _
                                    & errorSheet.Cells(j, 3) & "�v��" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "�ɂ��Ă��������B(" & .Cells(i, dataTypeColumn).Value & ")"
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
                        End If
                        
                    ElseIf condition = "�K�{" Then
                        If Not .Cells(i, dataTypeColumn).Value <> "" Then
                            '����G���[
                            logSheet.Cells(logWriteLine, 1) = i & "�s�ڂ́u" & .Cells(i, 3) & "�v���ڂ́u" _
                                    & errorSheet.Cells(j, 3) & "�v��" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "�ɂ��Ă��������B(" & .Cells(i, dataTypeColumn).Value & ")"
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
                        End If
                    End If
                ElseIf .Cells(i, 2).Value = "�Z" And errorSheet.Cells(j, 2).Value = "-" Then
                    If errorSheet.Cells(j, 5) = "���s����" And condition = "�܂܂Ȃ�" Then
                        If .Cells(i, dataTypeColumn).Value Like "*" & vbLf & "*" _
                        Or .Cells(i, dataTypeColumn).Value Like "*" & vbCrLf & "*" _
                        Or .Cells(i, dataTypeColumn).Value Like "*" & vbCr & "*" Then
                            
                            logSheet.Cells(logWriteLine, 1) = i & "�s�ڂ́u" & .Cells(i, 3) & "�v���ڂ́u" _
                                    & errorSheet.Cells(j, 3) & "�v��" & errorSheet.Cells(j, 5) _
                                    & errorSheet.Cells(j, 6) & "�ɂ��Ă��������B(" & .Cells(i, dataTypeColumn).Value & ")"
                            
                            logSheet.Cells(logWriteLine, 1).WrapText = False
                            logWriteLine = logWriteLine + 1
                            .Cells(i, dataTypeColumn).Interior.Color = RGB(255, 255, 0)
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
    End If
    
End Sub
Function init()
    Set itemSheet = Sheets(ITEM_SHEET)
    Set itemMetaSheet = Sheets(ITEM_META_SHEET)
    Set errorSheet = Sheets("Error��`")
    Set logSheet = Sheets("log")
    lastRow = itemSheet.Cells(5, 1).End(xlDown).row
    lastColumn = itemSheet.Cells(4, 1).End(xlToRight).column
    logSheet.Cells.Clear
    logWriteLine = 2
End Function
Function repaint()

    For i = 5 To lastRow
        With itemSheet
            dataType = .Cells(i, 7).Value
            dataType = IIf(.Cells(i, 8).Value = "�Z", "(����)" & dataType, dataType)
            
            If .Cells(i, 2).Value = "�~" Then
                '���ׂăO���[�ɓh��
                .Range(Cells(i, 1), Cells(i, 38)).Interior.Color = RGB(191, 191, 191)
            Else
                '�F����
                '�������s�̃f�[�^�^���`�V�[�g����T��
                For j = 4 To 31
                    If dataType = itemMetaSheet.Cells(2, j).Value Then
                        dataTypeColumn = j
                        Exit For
                    End If
                Next
                    
                '�f�[�^�^�ɉ����ċL�ڕs�v�ȏ����O���[�ɓh�F����
                For j = 3 To 37
                    colorChangeColumn = itemMetaSheet.Cells(j, 2).Value
                    If itemMetaSheet.Cells(j, dataTypeColumn).Value = "�Z" And colorChangeColumn > 0 Then
                        .Cells(i, colorChangeColumn).Interior.Color = NO_COLOR
                    ElseIf colorChangeColumn > 0 Then
                        .Cells(i, colorChangeColumn).Interior.Color = RGB(191, 191, 191)
                        .Cells(i, colorChangeColumn).Value = ""
                    End If
                Next
            End If
        End With
    Next

End Function
Sub a()
    Debug.Print Selection.Interior.Color
End Sub
Sub currentColorCheck()
    Dim currentColorCode As Long: currentColorCode = Selection.Interior.Color
    Dim Red As Integer: Red = currentColorCode Mod 256
    Dim Green As Integer: Green = Int(currentColorCode / 256) Mod 256
    Dim Blue As Integer: Blue = Int(currentColorCode / 256 / 256)
    
    Debug.Print "�F�l�F" & currentColorCode
    Debug.Print "�ԁF" & Red
    Debug.Print "�΁F" & Green
    Debug.Print "�F" & Blue
    Debug.Print "RGB(" & Red & "," & Green; "," & Blue & ")"
End Sub
