Attribute VB_Name = "subProcedure"
'###############################
'�@�\���F�h�L�������g�`�F�b�J�[1
'Author�Fokayasu jun
'�쐬���F2022/12/13
'�X�V���F2022/12/13
'COMMENT�F
'###############################

'�����s����
'�V�[�g�w��
Dim exeSheet As String
'#�`�F�b�N����
Sub check()
    '���s
    Call main
    '�����ʒm
    MsgBox noticeCount & "���`�F�b�N���܂����B�����b�Z�[�W�͍čl"
End Sub
'#���C������
Function main()
    '���ʏ����������iModule2�j
    Set customSettingCurrentRange = initialize
    
    '�Ǝ������������iModule1�j
    Call initializeInCustom(customSettingCurrentRange)
    
    '�t�@�C�����ƂɃ`�F�b�N���s
    Call scanLoopWithFile(srcDirPath)
        
    '���s����
    noticeCount = logWriteLine - 2
    
    '�I������
    Call finally
End Function
'#�@�\�Ǝ��̏������������s���B�����͋@�\�Z�N�V�����̍ŏ��̍��ڃ��x���̏ꏊ���w���Z��
Function initializeInCustom(customSettingCurrentRange As Range)

    '�����Ώۂ̃V�[�g���
    exeSheet = customSettingCurrentRange.Offset(0, 1).Value
    
    '�`�F�b�N���̊J�n�Z�����擾
    Set customSettingCurrentRange = getBottomEndRange(customSettingCurrentRange, 1)
    
    
    '�`�F�b�N�����擾
    Dim exe1StartRange As Range
    Set exe1StartRange = customSettingCurrentRange.Offset(0, 1)
    
    '�`�F�b�N���̃Z���͈͂��擾
    Dim exe1ArrayRange As Range
    Dim rightTimes As Integer: rightTimes = Cells(exe1StartRange.Offset(-1, 0).row, Columns.count).End(xlToLeft).Column / 2 - 2
    Set exe1ArrayRange = topSheet.Range(exe1StartRange, regionEndRange(exe1StartRange, headerFlag:=True, rightTimes:=rightTimes))
    
    '�Z���F�������擾���邽�߂̈ꎞ�����i�F�̒l�ݒ�j
    Dim editRange As Range
    Set editRange = exe1StartRange.Offset(0, 5)
    For i = exe1StartRange.row To getBottomEndRange(exe1StartRange, 1).row
        If topSheet.Cells(i, 9).Interior.Color <> 16777215 Or topSheet.Cells(i, 4) = "�Z���w�i�F" Then
            topSheet.Cells(i, 11).Value = topSheet.Cells(i, 9).Interior.Color
        ElseIf topSheet.Cells(i, 4) = "�Z�������F�i�S���j" _
            Or topSheet.Cells(i, 4) = "�Z�������F�i�ꕔ�j" _
            Or topSheet.Cells(i, 4) = "�w�b�_�����F" Then
            topSheet.Cells(i, 11).Value = topSheet.Cells(i, 9).Font.Color
        End If
    Next
    
    '�`�F�b�N����z��Ɋi�[
    exe1Array = IIf(exe1StartRange.Value = "", exe1StartRange, exe1ArrayRange)
    
    
    '�Z���F�������擾���邽�߂̈ꎞ�����i�ꎞ�l�N���A�j
'    For i = exe1StartRange.Row To getBottomEndRange(exe1StartRange, 1).Row
'        If topSheet.Cells(i, 11).Value = topSheet.Cells(i, 11).Interior.Color Then
'            topSheet.Cells(i, 11).Value = ""
'        End If
'    Next
    
    '�擾�n�`�F�b�N�p�B�܂��g�����낤�����U�c���Ă����B
'    For i = LBound(exe1Array) To UBound(exe1Array)
'        Debug.Print exe1Array(i, 1) & "�F" & exe1Array(i, 4) & "�F" & exe1Array(i, 6) & "�F" & exe1Array(i, 8)
'    Next
    
    Call initializeLogHeaderSetting
    
End Function
'#���O�̐ݒ�
Function initializeLogHeaderSetting()
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "�t�H���_"
    logSheet.Cells(1, 3) = "�t�@�C����"
    logSheet.Cells(1, 4) = "�V�[�g��"
    logSheet.Cells(1, 5) = "�`�F�b�N���"
    logSheet.Cells(1, 6) = "�T���s�E��ԍ�"
    logSheet.Cells(1, 7) = "���Ғl"
    logSheet.Cells(1, 8) = "�⏕���"
    logSheet.Cells(1, 9) = "���ʒl"
    logSheet.Cells(1, 10) = "�`�F�b�N����"
    logSheet.Cells(1, 11) = "�G���[���"
    logSheet.Cells(1, 12) = "����"
    logSheet.Cells(1, 13) = "���l"
End Function
'#�Ώۂ̑S�t�@�C���𑖍�����B�I�v�V�����ɉ����čċA�������s���B
Function scanLoopWithFile(argDirPath As String)
    '�t�H���_���̍ŏ��̃t�@�C�������擾
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '�������t�@�C����������ʉ߂��邩�ǂ���
        If isPassFile(argDirPath, currentFileName) Then
            Call exeCheck(argDirPath, currentFileName)
        End If
        
        '���̃t�@�C���������o���i�Ȃ���΃u�����N�j
        currentFileName = Dir()
    Loop
    
    If recursiveFlag Then
        '�t�H���_���̃T�u�t�H���_�����Ɏ擾
        For Each directory In objFSO.getfolder(argDirPath).SubFolders
            '�ċA����
            Call scanLoopWithFile(directory.Path)
        Next
    End If
End Function
'#
Function exeCheck(argDirPath As String, argFileName As String)
    
    Dim filePath As String: filePath = argDirPath & "\" & argFileName
    
    If isExcel(filePath) Then
        Call scanWithAllSheets(filePath, exeSheet)
    End If
    
End Function

'�V�[�g���`�F�b�N����
Function customProcess(ws As Worksheet)
    Dim checkType As String
    Dim checkLine As Integer
    Dim expectedValue As String
    Dim assistValue As String

    Dim tempArray As Variant
    Dim tempvalue As Variant
    Dim actualValue As String
    Dim result As String
    Dim errorCount As Integer
    Dim errorLog As String
    Dim headerRowNo As Integer
    
    If IsEmpty(exe1Array) Then
        Exit Function
    End If
    
    Dim latestRow As Integer
    latestRow = ws.Cells(Rows.count, 1).End(xlUp).row
    
    '�`�F�b�N���e�z��𑖍�����
    For i = LBound(exe1Array) To UBound(exe1Array)
        checkType = exe1Array(i, 1) '�`�F�b�N���
        '�T���s/��ԍ�
        checkLine = CInt(exe1Array(i, 4))
        '���Ғl
        expectedValue = exe1Array(i, 6)
        '�⏕���
        assistValue = exe1Array(i, 8)
        result = "��"
        errorLog = ""
        
        '�����A���g���Ȃ�������continue�ɂƂ΂��Ă�������
        
        If checkType = "�w�b�_�s�擾" Then
            '���w�b�_�s�̏ꏊ�������������`�F�b�N����i����A�擾�ł��Ȃ�������㑱������߂��������������ȁj
            For j = 1 To 20
                If ws.Cells(j, checkLine).Interior.Color = assistValue Then
                    actualValue = j
                    headerRowNo = j
                    Exit For
                End If
            Next
            '���ʋL�^�p
            result = IIf(actualValue <> expectedValue, "�~", result)
            If headerRowNo = 0 Then
                errorLog = "�w�b�_�s��������Ȃ��B"
            End If
            
        ElseIf checkType = "�w�b�_�i���ځj" Then
            '���w�b�_�̍��ږ�����e�E���ԂƂ��ɐ��������`�F�b�N����
            tempArray = Split(expectedValue, assistValue)
            For j = checkLine To ws.Cells(headerRowNo, Columns.count).End(xlToLeft).Column
                actualValue = actualValue & ws.Cells(headerRowNo, j) & assistValue
            Next
            actualValue = deleteEndText(actualValue)
            '���ʋL�^�p
            result = IIf(actualValue <> expectedValue, "�~", result)

        ElseIf checkType = "�w�b�_�����F" Then
            '���w�b�_�̕����F�𒲂ׂ�i�ꕔ�ɐF�����Ă�����̂͌��m���Ȃ��j
            For j = checkLine To ws.Cells(headerRowNo, Columns.count).End(xlToLeft).Column
                If ws.Cells(headerRowNo, j).Font.Color <> assistValue Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(headerRowNo, j).Address & "],"
                End If
            Next
            
        ElseIf checkType = "���͋K���i���X�g�j" Then
            '�������Ώےl�����X�g���ɂ��邩�ǂ������`�F�b�N����
            tempArray = Split(expectedValue, ",")
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And Not isExistArray(tempArray, ws.Cells(j, checkLine)) Then '�G���[�̏�����m�点�邱��
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "���͋K���i���K�\���j" Then
            '�������Ώےl�����K�\���Ƀ}�b�`���邩�ǂ������`�F�b�N����
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And Not isRegexpHit(ws.Cells(j, checkLine).Value, expectedValue) Then '�G���[�̏�����m�点�邱��
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "���͋K���i�ړ����j" Then
            '�������Ώےl���w��n�Ŏn�܂邩�ǂ������`�F�b�N����
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And Not isStartText(ws.Cells(j, checkLine), expectedValue) Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "���͋K���i�ڔ����j" Then
            '�������Ώےl���w��n�ŏI��邩�ǂ������`�F�b�N����
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And (Not isEndText(ws.Cells(j, checkLine), expectedValue)) Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "���͋K���i�ܕ����j" Then
            '�������Ώےl���w��n���܂ނ��ǂ������`�F�b�N����
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And Not InStr(ws.Cells(j, checkLine), expectedValue) > 0 Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
            Next
        ElseIf checkType = "�֎~����" Then
            '�������Ώےl�����X�g���ɂȂ����ǂ������`�F�b�N����
            tempArray = Split(expectedValue, ",")
            For j = headerRowNo + 1 To latestRow
                For k = LBound(tempArray) To UBound(tempArray)
                    tempArray(k) = IIf(tempArray(k) = "���s", vbLf, tempArray(k))
                    If withCheck(exe1Array, CInt(i), ws, CInt(j)) And InStr(ws.Cells(j, checkLine), tempArray(k)) > 0 Then
                        errorCount = errorCount + 1
                        errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                    End If
                Next
            Next
        ElseIf checkType = "�d���֎~" Then
            '���l�̏d�����Ȃ������`�F�b�N����
            For j = headerRowNo + 1 To latestRow
                For k = ws.Cells(Rows.count, 1).End(xlUp).row To j + 1 Step -1
                    If withCheck(exe1Array, CInt(i), ws, CInt(j)) And ws.Cells(j, checkLine) = ws.Cells(k, checkLine) Then
                        errorCount = errorCount + 1
                        errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                    End If
                Next
            Next
            
        ElseIf checkType = "�A����" Then
            '���l���A�����Ă��邩���`�F�b�N����
            tempNo = expectedValue
            Dim loopCount As Integer: loopCount = 0
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And tempNo + loopCount <> ws.Cells(j, checkLine) Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & ":" & ws.Cells(j, checkLine).Value & "],"
                End If
                loopCount = loopCount + 1
            Next
            
        ElseIf checkType = "�̍فi�ڐ����j" Then
            '���ڐ����̐ݒ��Ԃ��m�F����
            If expectedValue = "��\��" And (Not ActiveWindow.DisplayGridlines) Then
                '���Ғl�E���ʒl�Ƃ��ɔ�\��
                actualValue = "��\��"
            ElseIf expectedValue = "�\��" And ActiveWindow.DisplayGridlines Then
                '���Ғl�E���ʒl�Ƃ��ɕ\��
                actualValue = "�\��"
            Else
                errorCount = 1
            End If
            actualValue = IIf(ActiveWindow.DisplayGridlines, "�\��", "��\��")
            
        ElseIf checkType = "�̍فi�k�ځj" Then
            '���k�ڂ��`�F�b�N����
            If expectedValue <> ActiveWindow.Zoom Then
                errorCount = 1
            End If
            actualValue = ActiveWindow.Zoom
            
        ElseIf checkType = "�̍فi����޳�g�Œ�j" Then
            '���E�B���h�E�g�̌Œ肪�ݒ肳��Ă��邩�ǂ������`�F�b�N����
            If expectedValue = "���ݒ�" And (Not ActiveWindow.FreezePanes) Then
                '���Ғl�E���ʒl�Ƃ��ɖ��ݒ�
                actualValue = "���ݒ�"
            ElseIf expectedValue = "�ݒ�" And ActiveWindow.FreezePanes Then
                '���Ғl�E���ʒl�Ƃ��ɐݒ�
                actualValue = "�ݒ�"
            Else
                errorCount = 1
            End If
            actualValue = IIf(ActiveWindow.FreezePanes, "�ݒ�", "���ݒ�")
            
        ElseIf checkType = "�I���Z���ʒu" Then
            '���A�N�e�B�u�Z���̈ʒu�E�͈͂𒲂ׂ�
            If expectedValue <> Replace(Selection.Address, "$", "") Then
                errorCount = 1
            End If
            actualValue = Replace(Selection.Address, "$", "")
                
        ElseIf checkType = "�g�p�͈�" Then
            '��UsedRange�𒲂ׂ�
            If expectedValue <> Replace(ws.UsedRange.Address, "$", "") Then
                errorCount = 1
            End If
            actualValue = Replace(ws.UsedRange.Address, "$", "")
            
        ElseIf checkType = "�V�F�C�v�i���j" Then
            '���E�B���h�E�g�̌Œ肪�ݒ肳��Ă��邩�ǂ������`�F�b�N����
            If expectedValue <> ws.Shapes.count Then
                errorCount = 1
            End If
            actualValue = ws.Shapes.count
        
        ElseIf checkType = "�Z���w�i�F" Then
            '���Z���w�i�F�𒲂ׂ�
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And ws.Cells(j, checkLine).Interior.Color <> assistValue Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & "],"
                End If
            Next

        ElseIf checkType = "�Z�������F�i�S���j" Then
            '���Z���̕����F�𒲂ׂ�i�ꕔ�ɐF�����Ă�����̂͌��m���Ȃ��j
            For j = headerRowNo + 1 To latestRow
                If withCheck(exe1Array, CInt(i), ws, CInt(j)) And ws.Cells(j, checkLine).Font.Color <> assistValue Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "[" & ws.Name & ":" & ws.Cells(j, checkLine).Address & "],"
                End If
            Next
            
        ElseIf checkType = "" Then
        ElseIf checkType = "" Then
        End If
        
        If errorCount > 0 Then
            result = "�~"
        End If
        '�Ō�̃J���}���폜����
        errorLog = deleteEndText(errorLog)
        
        logSheet.Cells(logWriteLine, 1) = logWriteLine - 1 'No.
        logSheet.Cells(logWriteLine, 2) = ws.Parent.Path '�t�H���_
        logSheet.Cells(logWriteLine, 3) = ws.Parent.Name '�t�@�C����
        logSheet.Cells(logWriteLine, 4) = ws.Name '�V�[�g��
        logSheet.Cells(logWriteLine, 5) = checkType '�`�F�b�N���
        logSheet.Cells(logWriteLine, 6) = checkLine '�T���s/��ԍ�
        logSheet.Cells(logWriteLine, 7) = expectedValue '���Ғl
        logSheet.Cells(logWriteLine, 7).WrapText = False
        logSheet.Cells(logWriteLine, 8) = assistValue '�⏕���
        logSheet.Cells(logWriteLine, 8).WrapText = False
        logSheet.Cells(logWriteLine, 9) = actualValue '���ʒl
        logSheet.Cells(logWriteLine, 10) = result '�`�F�b�N����
        logSheet.Cells(logWriteLine, 11) = errorLog '�G���[���
        logSheet.Cells(logWriteLine, 11).WrapText = False
        time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
        logSheet.Cells(logWriteLine, 12) = time '����
        logSheet.Cells(logWriteLine, 13) = errorCount '�G���[��
        logWriteLine = logWriteLine + 1
        '������
        actualValue = ""
        result = ""
        errorCount = 0
        
        '�w�b�_���Ȃ������珈�����I������
        If headerRowNo = 0 Then
            Exit Function
        End If
    Next
End Function
'�`�F�b�N���̕t�я���
'�����F�`�F�b�N�Ώۂ��ǂ����������ɏ]���m�F����
'�����FcheckArray:�`�F�b�N�z��AcheckNo:�������̍s�ԍ��Aws:�`�F�b�N�Ώۂ̃V�[�g�Arow:�`�F�b�N�ΏۃV�[�g���̍s�ԍ�
'�ߒl�F�^�U�l�itrue:�`�F�b�N�ΏہAfalse:�`�F�b�N�Ώۂł͂Ȃ��j
Function withCheck(checkArray As Variant, checkNo As Integer, ws As Worksheet, row As Integer)

    Dim checkColumn As Integer
    Dim checkValue As String
    Dim checkType As String
    Dim checkConnector As String
    Dim currentResult As Boolean
    Dim checkedValue As String
    Dim totalResult As Boolean: totalResult = True
    
    Dim checkLogic As String
    Dim checkReturnArray As Variant
    
    Dim collection1 As MatchCollection
    Dim collection2 As MatchCollection
    Dim collection3 As MatchCollection
    Dim collection4 As MatchCollection
    Dim collection5 As MatchCollection
    Dim andor As String
    
    checkLogic = checkArray(checkNo, 9)
    '�������W�b�N�̗L���ŕ]�����@��ς���i���W�b�N���A�����珇�����j
    If checkLogic <> "" Then
        '�e�]������y�����̔z��ɕϊ�����
        checkReturnArray = convertArrayFrom2dTo2d(checkArray, checkNo, 10)
        '�������W�b�N���犇�ʒP�ʂ̃R���N�V�������擾
        '��F�u(1 or 2) and (3 or 4 or 5)�v����u(1 or 2)�v �Ɓu(3 or 4 or 5)�v�𒊏o�j
        Set collection1 = regexpHitCollection(checkLogic, "\(.+?\)")
        '�擾�ł������ʖ��ɒ��̏�����]������
        For Each hit1 In collection1
            '�������W�b�N�̐��l�����𒊏o����
            Set collection2 = regexpHitCollection(hit1.Value, "\d")
            '1 2 3 4 5 �𒊏o
            For Each hit2 In collection2
                '���[�v���̐��l�����������擾����
                checkColumn = CInt(checkReturnArray(CInt(hit2.Value), 1))
                
                If checkColumn <> 0 Then
                    '�����l
                    checkValue = checkReturnArray(CInt(hit2.Value), 2)
                    '�������
                    checkType = checkReturnArray(CInt(hit2.Value), 3)
                    '���������l
                    checkedValue = ws.Cells(row, checkColumn).Value
                
                    currentResult = getCheckReuslt(checkType, checkedValue, checkValue)
                
                    '���ʓ��̐��l�𓱏o�����^�U�l�Œu��
                    checkLogic = Replace(checkLogic, hit2.Value, currentResult)
                End If
            Next
        Next
        '�������܂łŐ^�U�l�ϊ���̕����񂪒a��
        Debug.Print "�@" & checkLogic
        '���ʖ��̘_�����𓝍�����
        Set collection3 = regexpHitCollection(checkLogic, "\(.+?\)")
        For Each hit1 In collection3
            '���ʓ��̒l���R���N�V�����Ŏ擾
            Set collection4 = regexpHitCollection(hit1.Value, "[a-zA-Z]{2,5}")
            For i = 0 To collection4.count - 1 Step 2
                If i = 0 Then
                    '�ŏ��̏������،��ʂ͂��̂܂܏����l�Ƃ��Đݒ肷��
                    totalResult = CBool(collection4.Item(i))
                Else
                    '2�ڈȍ~�̏�����AND/OR�ƑO��̌��ʂɂ���ē��o����B������-1�͑O��̏�����ʂ��擾����������
                    currentResult = CBool(collection4.Item(i))
                    andor = CStr(collection4.Item(i - 1))
                    totalResult = deriveTotalResult(totalResult, currentResult, andor)
                End If
            Next
            '���ʂ𓱏o�����^�U�l�Œu��
            checkLogic = Replace(checkLogic, hit1.Value, totalResult)
        Next
        Debug.Print "�A" & checkLogic
        
        '���ʓ��m�̐^�U�l��
        Set collection5 = regexpHitCollection(checkLogic, "[a-zA-Z]{2,5}")
        For i = 0 To collection5.count - 1 Step 2
            If i = 0 Then
                '�ŏ��̏������،��ʂ͂��̂܂܏����l�Ƃ��Đݒ肷��
                totalResult = CBool(collection5.Item(i))
            Else
                '2�ڈȍ~�̏�����AND/OR�ƑO��̌��ʂɂ���ē��o����B������-1�͑O��̏�����ʂ��擾����������
                currentResult = CBool(collection5.Item(i))
                andor = CStr(collection5.Item(i - 1))
                totalResult = deriveTotalResult(totalResult, currentResult, andor)
            End If
        Next
        Debug.Print "�B" & totalResult
    Else

        '10�������̊J�n�ʒu
        For i = 10 To UBound(checkArray, 2) Step 8
            '������i�u�����N�Ȃ�0�ɂȂ�j
            checkColumn = CInt(checkArray(checkNo, i))
            
            
            If checkColumn <> 0 Then
                '�T���l
                checkValue = checkArray(checkNo, i + 2)
                '�������
                checkType = checkArray(checkNo, i + 4)
                '�_�������q
                checkConnector = checkArray(checkNo, i - 2)
                '���������l
                checkedValue = ws.Cells(row, checkColumn).Value
        
                currentResult = getCheckReuslt(checkType, checkedValue, checkValue)
        
                If i = 10 Then
                    '�ŏ��̏������،��ʂ͂��̂܂܏����l�Ƃ��Đݒ肷��
                    totalResult = currentResult
                Else
                    '���݂̍ŐV�^�U�l�ƍ���̃`�F�b�N���ʂ��g���_�������q�œ��o����
                    totalResult = deriveTotalResult(totalResult, currentResult, checkConnector)
                End If
            End If
        Next
    End If
    withCheck = totalResult
End Function

'
'�����F2�����z���2�����z���
'�����F
'�ߒl�F2�����z��
Function convertArrayFrom2dTo2d(srcArray As Variant, argReadRow As Integer, argReadStartColumn As Integer)
    Dim returnArray() As Variant
    '�v�f���̐ݒ�
    ReDim returnArray(UBound(srcArray, 2), 3)
    Dim count As Integer: count = 1
    
    '�񎟌��ڂɂ��ă��[�v���s��
    For i = argReadStartColumn To UBound(srcArray, 2) Step 8

        If srcArray(argReadRow, i) <> "" Then
            returnArray(count, 0) = count
            '������
            returnArray(count, 1) = srcArray(argReadRow, i)
            '�����l
            returnArray(count, 2) = srcArray(argReadRow, i + 2)
            '�������
            returnArray(count, 3) = srcArray(argReadRow, i + 4)
            count = count + 1
        End If
    Next
    convertArrayFrom2dTo2d = returnArray
End Function
'�����F
'�����F
'�ߒl�F
Function getCheckReuslt(checkType As String, checkedValue As String, checkValue As String)

    Dim result As Boolean
    
    If checkType = "���L�̒l�ƈ�v����" Then
        result = (checkedValue = checkValue)
    
    ElseIf checkType = "���L�̒l���܂�" Then
        result = InStr(checkedValue, checkValue) > 0
    
    ElseIf checkType = "���L�̒l���܂܂Ȃ�" Then
        result = InStr(checkedValue, checkValue) = 0
    
    ElseIf checkType = "���L�̒l�Ŏn�܂�" Then
        result = isStartText(checkedValue, checkValue)
    
    ElseIf checkType = "���L�̒l�ŏI���" Then
        result = isEndText(checkedValue, checkValue)
    
    End If
    
    getCheckReuslt = result
    
End Function
