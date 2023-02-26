Attribute VB_Name = "common"
'����{���
'���s�V�[�g
Public topSheet As Worksheet
'���s���O�V�[�g
Public logSheet As Worksheet
'�t�@�C������I�u�W�F�N�g
Public objFSO As Object
'���t�H���_�p�X
Public srcDirPath As String
'��t�H���_�p�X
Public distDirPath As String
'�����ʃI�v�V����
'�ċA�����t���O
Public recursiveFlag As Boolean
'���O�L�^�t���O
Public logFlag As Boolean
'�Ώۃt�@�C���g���q
Public targetFileExtArray As Variant
'�X�V�����iFROM�j
Public lastUpdateDateFrom As Date
'�X�V�����iTO�j
Public lastUpdateDateTo As Date
'�t�H���_�\���Č��t���O
Public dirLevelCopyFlag As Boolean
'�Ώۃt�@�C�������Ɋg���q���ΏۂƂ���t���O
Public extensionUseFlag As Boolean
'�Ώۃt�@�C������
Public targetFilterArray As Variant
'�����̑�
'���O�L�^�s
Public logWriteLine As Integer
'���s����
Public noticeCount As Integer
'���ԋL�^�p
Public time As String
'�@�\�Ǝ��̐ݒ荀�ڊJ�n�ʒu
Public customSettingCurrentRange As Range
'���@�\�Ǝ��̕ϐ�
'�`�F�b�N�����i�[�z��
Public exe1Array As Variant
'�w�b�_�s
Public headerRowNo As Integer
'#�����F��������
'#�����F�Ȃ�
'#�ߒl�FRange:�@�\�Z�N�V�����̍ŏ��̍��ڃ��x���Z��
Function initialize()

    '�����s�����
    '���s�V�[�g
    Set topSheet = ThisWorkbook.Sheets(1)
    '���O�V�[�g
    If isExistCheckToSheet(ActiveWorkbook, "log") Then
        Set logSheet = ThisWorkbook.Sheets(2)
    Else
        Set logSheet = Sheets.Add(After:=Sheets(1))
        logSheet.Name = "log"
    End If
    
    '�t�@�C������I�u�W�F�N�g
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    With topSheet
        '����{���
        Dim baseSectionInputStartRange As Range
        Set baseSectionInputStartRange = getBottomEndRange(.Cells(1, 3), 1)
        
        '�R�s�[���t�H���_
        srcDirPath = baseSectionInputStartRange.Offset(0, 1).value
        
        '�R�s�[��t�H���_
        distDirPath = baseSectionInputStartRange.Offset(1, 1).value

        '�����ʃI�v�V����
        Dim commonOptionSectionInputStartRange As Range
        Set commonOptionSectionInputStartRange = getBottomEndRange(baseSectionInputStartRange, 2)
        
        '�ċA�����t���O
        recursiveFlag = commonOptionSectionInputStartRange.Offset(0, 1).value = "����"
        
        '���O�o�͐ݒ�
        logFlag = commonOptionSectionInputStartRange.Offset(1, 1).value = "����"
        
        '�Ώۃt�@�C���`��
        targetFileExtArray = Split(commonOptionSectionInputStartRange.Offset(2, 1).value, ",")
        
        '�X�V�����iFROM�j
        lastUpdateDateFrom = commonOptionSectionInputStartRange.Offset(3, 1).value
        
        '�X�V�����iTO�j
        lastUpdateDateTo = commonOptionSectionInputStartRange.Offset(4, 1).value
        
        '�t�H���_�\���Č��t���O
        dirLevelCopyFlag = commonOptionSectionInputStartRange.Offset(5, 1).value = "����"
        
        '�g���q���p�t���O
        extensionUseFlag = commonOptionSectionInputStartRange.Offset(6, 1).value = "����"
        
        '�����i1:�l�A2,4,5,:�s�g�p�A3:������ʁA6:AND/OR�j
        Dim targetFilterStartRange As Range
        Set targetFilterStartRange = commonOptionSectionInputStartRange.Offset(7, 1)
        Dim targetFilterEndRange As Range
        Set targetFilterEndRange = regionEndRange(targetFilterStartRange, rightTimes:=2)
        targetFilterArray = IIf(targetFilterStartRange.value = "", targetFilterStartRange, .Range(targetFilterStartRange, targetFilterEndRange))

        '���@�\�Z�N�V�����̍ŏ��̃��x���Z����Ԃ�
        Set initialize = getBottomEndRange(commonOptionSectionInputStartRange, 2)

    End With
    
    '���O�̏����o���s
    logWriteLine = 2
    
    '�L�^����
    time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
    
    '�����s�������
    '���s���̉�ʕ`���Â��ɂ���
    Application.DisplayAlerts = False
    '�g���K�[���̎����}�N�����N�������Ȃ�
    Application.EnableEvents = False
    '��ʒ�~
    Application.ScreenUpdating = False
End Function
'#�����F�I������
'#�����F�Ȃ�
'#�ߒl�F�Ȃ�
Function finally()
    '���s���̉�ʕ`������ɖ߂�
    Application.DisplayAlerts = True
    '�g���K�[�������}�N�����������
    Application.EnableEvents = True
    '��ʒ�~����
    Application.ScreenUpdating = True
End Function
'#�����F�w�肳�ꂽ�V�[�g�̃Z������w��񐔕����ړ������Z����Ԃ�
'#�����FargRange:�J�n�ʒu�Z���AargTimes:�ړ���
'#�ߒl�F�ړ���̃Z�����
Function getBottomEndRange(argRange As Range, argTimes As Integer)
    Dim returnRange As Range
    Set returnRange = argRange
    For i = 1 To argTimes
        Set returnRange = topSheet.Range(returnRange.Address).End(xlDown)
    Next
    '�߂�l
    Set getBottomEndRange = returnRange
End Function
'#�����F�w�肳�ꂽ�V�[�g�̃Z������w��񐔕��E�ړ������Z����Ԃ�
'#�����FargRange:�J�n�ʒu�Z���AargTimes:�ړ���
'#�ߒl�F�ړ���̃Z�����
Function getRightEndRange(argRange As Range, argTimes As Integer)
    Dim returnRange As Range
    Set returnRange = argRange
    For i = 1 To argTimes
        Set returnRange = topSheet.Range(returnRange.Address).End(xlToRight)
    Next
    '�߂�l
    Set getRightEndRange = returnRange
End Function
'#�����F�w�肳�ꂽ�t�H���_���Ȃ���΍쐬����
'#�����FargDirPath:�t�H���_�p�X�i��΃p�X�j
'#�ߒl�F�Ȃ�
Function createDirectory(argDirPath As String)
    If Not objFSO.FolderExists(argDirPath) Then
        objFSO.CreateFolder (argDirPath)
    End If
End Function
'#�����F�������t�@�C����������ʉ߂��邩�ǂ�����Ԃ��i�������̓O���[�o���ϐ�����Ƃ�j
'#�����FargDirPath:�t�H���_�p�X�i��΃p�X�j�AargFileName:�t�@�C����
'#�ߒl�F�^�U�l�itrue:�ʉ߁Afalse:�s�K���j
Function isPassFile(argDirPath As String, argFileName As String)
    '�t�@�C���t���p�X
    Dim filePath As String: filePath = argDirPath & "\" & argFileName
    '�g���q�̎擾
    Dim fileExt As String: fileExt = objFSO.GetExtensionName(filePath)
    '�g���q���Ȃ����t�@�C����
    Dim checkFileName As String: checkFileName = argFileName
    If Not extensionUseFlag Then
        checkFileName = Replace(argFileName, "." & fileExt, "")
    Else
    End If
    '�Ώۃt�@�C�������ɊY�����邩
    If isExistArray(targetFileExtArray, fileExt) _
        And isPassConditionCheck(checkFileName) _
        And isPassUpdateDate(filePath) Then
        '�߂�l
        isPassFile = True
        Exit Function
    End If
    '�߂�l
    isPassFile = False
End Function
'#�����F�w�肳�ꂽ�l���w�肳�ꂽ�z����ɑ��݂��邩�ǂ�����Ԃ�
'#�����FtargetArray:�����z��AcheckValue:���ؒl
'#�ߒl�F�^�U�l�itrue:����Afalse:�Ȃ��j
Function isExistArray(targetArray As Variant, checkValue As String)
    isExistArray = False
    'UBound�̖߂�l�F-1�͗v�f��0������
    If UBound(targetArray) = -1 Then
        isExistArray = True
        Exit Function
    End If
    
    For i = LBound(targetArray) To UBound(targetArray)
        If targetArray(i) = checkValue Then
            isExistArray = True
            Exit For
        End If
    Next
End Function
'#�����F�w�肳�ꂽ�t�@�C�������������A���̌��ʂ�Ԃ�
'#�����FtargetArray:�����z��AcheckValue:���ؒl
'#�ߒl�F�^�U�l�itrue:OK�Afalse:NO�j
Function isPassConditionCheck(argFileName As String)
    '�`�F�b�N�Ώۂ̒l
    Dim checkValue As String
    '���[�v�����ݎ���̌��،��ʐ^�U�l
    Dim currentResult As Boolean
    '�݌v�̌��،��ʐ^�U�l
    Dim totalResult As Boolean
    '���؏����̎��
    Dim conditionType As String
    '�����̎�ށBAnd��Or��
    Dim andor As String

    If IsEmpty(targetFilterArray) Then
        isPassConditionCheck = True
        Exit Function
    End If
    
    '�ŏ��̗v�f����̏ꍇ�A�����w��͂Ȃ��Ɣ��f���A���ׂ�true��Ԃ�
    If targetFilterArray(LBound(targetFilterArray, 1), 1) = "" Then
        isPassConditionCheck = True
        Exit Function
    End If
    
    '���[�v���ł��g�����߈�x�ϐ��ɂ����
    Dim minIndex As Integer: minIndex = LBound(targetFilterArray)
    For i = minIndex To UBound(targetFilterArray)
    
        checkValue = targetFilterArray(i, 1)
        conditionType = targetFilterArray(i, 3)
        
        If checkValue = "" Or conditionType = "" Then
            GoTo continue
        End If
        
        If conditionType = "���L�Ƀt�@�C��������v����" Then
            currentResult = (argFileName = checkValue)
            
        ElseIf conditionType = "���L���t�@�C�����Ɋ܂�" Then
            currentResult = InStr(argFileName, checkValue) > 0
            
        ElseIf conditionType = "���L���t�@�C�����Ɋ܂܂Ȃ�" Then
            currentResult = InStr(argFileName, checkValue) = 0
            
        ElseIf conditionType = "���L����t�@�C�������n�܂�" Then
            currentResult = isStartText(argFileName, checkValue)
            
        ElseIf conditionType = "���L�Ńt�@�C�������I���" Then
            currentResult = isEndText(argFileName, checkValue)
            
        ElseIf conditionType = "���L�̐��K�\���Ɉ�v����" Then
            currentResult = isRegexpHit(argFileName, checkValue)
            
        End If
        
        If i = minIndex Then
            '�ŏ��̏������،��ʂ͂��̂܂܏����l�Ƃ��Đݒ肷��
            totalResult = currentResult
        Else
            '2�ڈȍ~�̏�����AND/OR�ƑO��̌��ʂɂ���ē��o����B������-1�͑O��̏�����ʂ��擾����������
            andor = CStr(targetFilterArray(i - 1, 6))
            totalResult = deriveTotalResult(totalResult, currentResult, andor)
        End If
continue:
    Next
    isPassConditionCheck = totalResult
End Function
'#�����F�����̐^�U�l�ƐV���Ȑ^�U�l��������ɂ���Đ^�U�l�𓱏o����
'#�����FtotalResult:�����̐^�U�l�AcurrentResult:�V�����^�U�l�Aandor:������ʁi���E�܂��́j
'#�ߒl�F�^�U�l�itrue:OK�Afalse:NO�j
Function deriveTotalResult(totalResult As Boolean, currentResult As Boolean, andor As String)
    If andor = "����" Or LCase(andor) = "and" Then
        totalResult = totalResult And currentResult
    ElseIf andor = "�܂���" Or LCase(andor) = "or" Then
        totalResult = totalResult Or currentResult
    Else
        '���w�莞��OR�����Ƃ��Ĉ���
        totalResult = totalResult Or currentResult
    End If
    '�߂�l
    deriveTotalResult = totalResult
End Function
'#�����F�����񂪎w��̕����Ŏn�܂邩�ǂ�����Ԃ�
'#�����FlargeText:���ؑΏۂ̕�����AsearchText:���앶����
'#�ߒl�F�^�U�l�itrue:�n�܂�Afalse:�n�܂�Ȃ��j
Function isStartText(largeText As String, searchText As String)
    isStartsText = False
    If Len(searchText) > Len(largeText) Then
        '���؃e�L�X�g���팟�؃e�L�X�g�̒����𒴂���ꍇ�`�F�b�N���������Ȃ�
        Exit Function
    End If
  
    If Left(largeText, Len(searchText)) = searchText Then
        isStartText = True
    End If
End Function
'#�����F�����񂪎w��̕����ŏI��邩�ǂ�����Ԃ�
'#�����FlargeText:���ؑΏۂ̕�����AsearchText:���앶����
'#�ߒl�F�^�U�l�itrue:�I���Afalse:�I���Ȃ��j
Function isEndText(largeText As String, searchText As String)
    isEndText = False
    If Len(searchText) > Len(largeText) Then
        '���؃e�L�X�g���팟�؃e�L�X�g�̒����𒴂���ꍇ�`�F�b�N���������Ȃ�
        Exit Function
    End If

    If Right(largeText, Len(searchText)) = searchText Then
        isEndText = True
    End If
End Function
'#�����F�t�@�C���̍ŏI�X�V�����������ǂ�����Ԃ�
'#�����FargFilePath:�t�@�C���̃t���p�X
'#�ߒl�F�^�U�l�itrue:�������Afalse:�����O�j
Function isPassUpdateDate(argFilePath As String)
    isPassUpdateDate = True
    Dim fileUpdateDate As Date: fileUpdateDate = objFSO.getFile(argFilePath).DateLastModified
    '���w��̏ꍇ�u0�v�ɂȂ邽�߂��̏���
    If lastUpdateDateFrom <> 0 And Not lastUpdateDateFrom <= fileUpdateDate Then
        isPassUpdateDate = False
    End If
    If lastUpdateDateTo <> 0 And Not fileUpdateDate <= lastUpdateDateTo Then
        isPassUpdateDate = False
    End If
    
End Function
'#�����F�t�@�C���̋֎~�������폜����
'#�����FfileName:�t�@�C����
'#�ߒl�F�֎~�����폜��̕�����
Function replaceTabooStrWithFileName(fileName As String)
    Dim tabooStringArray As Variant: tabooStringArray = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each taboo In tabooStringArray
        fileName = Replace(fileName, taboo, "")
    Next
    replaceTabooStrWithFileName = fileName
End Function
'#�����FExcel�V�[�g�̋֎~�������폜����
'#�����FsheetName:�V�[�g��
'#�ߒl�F�֎~�����폜��̕�����
Function replaceTabooStrWithSheetName(sheetName As String)
    Dim tabooStringArray As Variant: tabooStringArray = Array(":", "�F", "\", "��", "?", "�H", "[", "�m", "]", "�n", "/", "�^", "*", "��")
    For Each taboo In tabooStringArray
        sheetName = Replace(sheetName, taboo, "")
    Next
    replaceTabooStrWithSheetName = sheetName
End Function
'#�����F�w�肳�ꂽ�l���w�肳�ꂽ�z����̉��Ԗڂɑ��݂��邩��Ԃ�
'#�����FtargetArray:�����z��AcheckValue:�����l
'#�ߒl�F�l���ŏ��ɏo���C���f�b�N�X
Function isExistArrayReturnIndex(targetArray As Variant, checkValue As String)
    isExistArrayReturnIndex = -1
    
    'UBound�̖߂�l�F-1�͗v�f��0�������B���̏ꍇ�A-1��Ԃ�
    If UBound(targetArray) = -1 Then
        isExistArrayReturnIndex = -1
        Exit Function
    End If
    
    For i = LBound(targetArray) To UBound(targetArray)
        If targetArray(i) = checkValue Then
            '�q�b�g�����C���f�b�N�X��߂�l�ɐݒ�
            isExistArrayReturnIndex = i
            Exit For
        End If
    Next
End Function
'#�����F�n���ꂽ�t�@�C���̕����R�[�h��Ԃ�
'#�����FfilePath:�t�@�C���̃t���p�X
'#�ߒl�F�����R�[�h�iSJIS�AUTF8�j
Function judgeFileCharSet(filePath As String)
    '����̂��߂Ƀo�C�i�����[�h�Ŏ擾����
    Dim bytCode() As Byte
    With CreateObject("ADODB.Stream")
        .Type = 1 '�o�C�i���ŊJ������
        .Open
        .LoadFromFile filePath
        bytCode = .read
        .Close
    End With
    judgeFileCharSet = judgeCode(bytCode)
End Function
'#�����F�񎟌��z������̗�i2�����ڂ̔z��j�ňꎟ���z��ɕϊ�����
'#�����FsrcArray:�񎟌��z��AargCol:�񎟌��ڂ̃C���f�b�N�X
'#�ߒl�F�ꎟ���z��
Function convertArrayFrom2dTo1d(srcArray As Variant, argCol As Integer)
    Dim returnArray As Variant
    For i = LBound(srcArray) To UBound(srcArray)
        If i = LBound(srcArray) Then
            '�ŏ��̗v�f�́uPreserve�v���g��Ȃ����炱�̕���
            ReDim returnArray(0)
            returnArray(0) = srcArray(i, argCol)
        Else
            '�z��̗v�f����1���₷
            ReDim Preserve returnArray(UBound(returnArray) + 1)
            '���₵���v�f�ɒl���i�[����
            returnArray(UBound(returnArray)) = srcArray(i, argCol)
        End If
    Next
    convertArrayFrom2dTo1d = returnArray
End Function
'#�����F�S�V�[�g�𑖍�����B�ڍ׏����͋@�\���Ŏ�������
'#�����Fwb:�u�b�N�AexeSheet:�����Ώۂ̃V�[�g�i�V�[�g��/�ԍ��B�J���}��؂�ŕ����w����j
'#�ߒl�F�Ȃ�
Function scanWithAllSheets(filePath As String, exeSheet As String)
    Dim wb As Workbook
    '�����Ńu�b�N���J��
    Set wb = Workbooks.Open(fileName:=filePath, UpdateLinks:=0)

    '�����p
    Dim ws As Worksheet
    Dim exeSheetNo As Integer
    Dim exeSheetName As String
    Dim exeSheetArray As Variant
    exeSheetArray = Split(exeSheet, ",")
    'TODO:�����uexeSheet�v�̒l�Ɂu�S�v�������Ă����̑I�����郋�[�v���g�킸�A�𒼂Ƀu�b�N�̂��ׂĂ�Ώۂɂ��郍�W�b�N�g����
    
    For i = LBound(exeSheetArray) To UBound(exeSheetArray)
        If IsNumeric(exeSheetArray(i)) Then
            '�V�[�g��ԍ��Ŏw��
            exeSheetNo = CInt(exeSheetArray(i))
            If isExistCheckToSheet(wb, exeSheetNo) Then
                Set ws = wb.Worksheets(exeSheetNo)
                '�Ǝ������ɃV�[�g��n��
                Call customProcess(ws)
            End If
        Else
            '�V�[�g�𖼑O�Ŏw��
            exeSheetName = CStr(exeSheetArray(i))
            If isExistCheckToSheet(wb, exeSheetName) Then
                Set ws = wb.Worksheets(exeSheetName)
                '�Ǝ������ɃV�[�g��n��
                Call customProcess(ws)
            End If
            
        End If
    Next
    
    wb.Close SaveChanges:=False
End Function
'#�����F�w�肳�ꂽ�V�[�g�̎g�p�͈͍ŏI�Z�����擾����
'#�����Fws:�V�[�g
'#�ߒl�F�g�p�͈͍ŏI�Z���̃A�h���X
Function usedLastRange(ws As Worksheet)
    Dim addressArray As Variant
    addressArray = Split(ws.UsedRange.Address, ":")
    If UBound(addressArray) = 0 Then
        '�P��Z���̏ꍇ
        Set lastRange = ws.Range(Split(ws.UsedRange.Address, ":")(0))
    Else
        '�����Z���ɓn���ꍇ
        Set lastRange = ws.Range(Split(ws.UsedRange.Address, ":")(1))
    End If
    Set usedLastRange = lastRange
End Function
'#�����F�w�肳�ꂽ�Z�����N�_�Ƃ����͈͂̉E���Z�����擾����B�E�ړ��Ƀw�b�_���g�����A�E�ړ�������s�������w��ł���
'#�����FstartRange:�J�n�ʒu�Z���AheaderFlag:�w�b�_�s�ŉE�ړ������邩�ArightTimes:�E�ړ���
'#�ߒl�F�n�_����擾�ł����E���̃Z��
Function regionEndRange(startRange As Range, Optional headerFlag As Boolean = False, Optional rightTimes As Integer = 1)
        Dim rightEndRange As Range
        Dim tempRange As Range
        Set tempRange = startRange
        '�w�b�_�t���O��TRUE�̏ꍇ�A�w�b�_�ŉE�����̍ŏI����擾����
        If headerFlag Then
            Set tempRange = startRange.Offset(-1, 0)
        End If
        
        Set rightEndRange = getRightEndRange(tempRange, rightTimes)
        Dim bottomEndRange As Range
        If getBottomEndRange(startRange, 1).value <> "" Then
            Set bottomEndRange = getBottomEndRange(startRange, 1)
        Else
            Set bottomEndRange = startRange
        End If
        '�����̊J�n�Z������s�����A������ɂ��炵���Z����߂�l�ɐݒ�
        Set regionEndRange = startRange.Offset(bottomEndRange.row - startRange.row, rightEndRange.Column - startRange.Column)
End Function
'#�����F�w�肳�ꂽ�t�@�C���̒��g���擾����
'#�����FfullPath:�t�@�C���̃t���p�X
'#�ߒl�F�e�L�X�g�t�@�C���̓��e
Function getFileText(fullPath As String)
    '�����R�[�h����
    Dim charset As String: charset = judgeFileCharSet(fullPath)
    If charset = "UTF8" Then
        charset = "UTF-8"
    ElseIf charset = "SJIS" Then
        charset = "SHIFT-JIS"
    End If
    
    With CreateObject("ADODB.Stream")
        .charset = charset
        .Open
        .LoadFromFile fullPath
        '�߂�l
        getFileText = .ReadText
        .Close
    End With
End Function
'#�����F�w�肳�ꂽ�p�X����t�H���_�������̓t�@�C�������擾����B
'#�����FargFilePath:�t�@�C���̃t���p�X�Awitch:1�Ȃ�t�H���_�A����ȊO�Ȃ�t�@�C����
'#�ߒl�F�t�H���_�p�X�������̓t�@�C����
Function extractDirOrFile(argFilePath As String, witch As Integer)
    Dim dirs As Variant: dirs = Split(argFilePath, "\")
    If witch = 1 Then
        dirAndFileFromFullPath = Left(argFilePath, Len(argFilePath) - Len(dirs(UBound(dirs))) - 1)
    Else
        dirAndFileFromFullPath = dirs(UBound(dirs))
    End If
End Function
'#�����F�w�肳�ꂽ�u�b�N�Ɏw�肳�ꂽ�V�[�g�����݂��邩�ǂ�����Ԃ�
'#�����Fwb:�����u�b�N�AcheckSheet:�T���V�[�g�i��/�ԍ��ǂ�����j
'#�ߒl�F�^�U�l�itrue:����Afalse:�Ȃ��j
Function isExistCheckToSheet(wb As Workbook, checkSheet As Variant)
    isExistCheckToSheet = False
    For Each ws In wb.Worksheets
        If Not IsNumeric(checkSheet) Then
            If ws.Name = checkSheet Then
                isExistCheckToSheet = True
            End If
        End If
    Next
    
    If IsNumeric(checkSheet) Then
        '�w��l�����l�̏ꍇ�A�S�V�[�g����菬�����������ǂ��������݃`�F�b�N�ƂȂ�
        isExistCheckToSheet = checkSheet <= wb.Worksheets.count
    End If
End Function
'#�����F�w�肳�ꂽ�l�����K�\���p�^�[���Ɉ�v���邩�ǂ��������؂���
'#�����Ftext:����������Apattern:���K�\���p�^�[��
'#�ߒl�F�^�U�l�itrue:��v����Afalse:��v���Ȃ��j
Function isRegexpHit(text As String, pattern As String)
    Set regexp = CreateObject("VBScript.RegExp")
    With regexp
         '�������鐳�K�\������
         .pattern = pattern
        '�啶���������̋�ʁiTrue�F���Ȃ��AFalse�F����j
        .IgnoreCase = False
        '������̍Ō�܂Ō����iTrue�F����AFalse�F���Ȃ��j
        .Global = True
        '�߂�l
        isRegexpHit = .test(text)
    End With
End Function
'#�����F�w�肳�ꂽ�l�̂������K�\���p�^�[���Ɉ�v������̂��R���N�V�����ŕԂ�
'#�����Ftext:����������Apattern:���K�\���p�^�[��
'#�ߒl�F�R���N�V�����i�Q�Ɛݒ�FMicrosoft VBScript Regular Expressions 5.5�j
Function regexpHitCollection(text As String, pattern As String)
    Set regexp = CreateObject("VBScript.RegExp")
    With regexp
         '�������鐳�K�\������
         .pattern = pattern
        '�啶���������̋�ʁiTrue�F���Ȃ��AFalse�F����j
        .IgnoreCase = False
        '������̍Ō�܂Ō����iTrue�F����AFalse�F���Ȃ��j
        .Global = True
        '�߂�l
        Set regexpHitCollection = .Execute(text)
    End With
End Function
'#�����F�w�肳�ꂽ�t�@�C���p�X���G�N�Z�����ǂ�����Ԃ�
'#�����FfilePath:�t�@�C���t���p�X
'#�ߒl�F�^�U�l�itrue:�G�N�Z���Afalse:�G�N�Z���ȊO�j
Function isExcel(filePath As String)
    isExcel = InStr(objFSO.getFile(filePath).Type, "Excel") > 0
End Function
'#�����F�w�肳�ꂽ��������������폜�����������Ԃ�
'#�����Ftext:������AdeleteLength:�폜���镶����
'#�ߒl�F�폜�㕶����
Public Function deleteStartText(text As String, Optional deleteLength As Long = 1) As String
    If Len(text) >= deleteLength Then
        deleteStartText = Right(text, Len(text) - deleteLength)
    Else
        deleteStartText = text
    End If
End Function
'#�����F�w�肳�ꂽ����������납��폜�����������Ԃ�
'#�����Ftext:������AdeleteLength:�폜���镶����
'#�ߒl�F�폜�㕶����
Public Function deleteEndText(text As String, Optional deleteLength As Long = 1) As String
    If Len(text) >= deleteLength Then
        deleteEndText = Left(text, Len(text) - deleteLength)
    Else
        deleteEndText = text
    End If
End Function
'#�����F�w�肳�ꂽ�t�H���_�̍ŏI�X�V�t�@�C������Ԃ�
'#�����FargDirPath:�t�H���_�p�X
'#�ߒl�F�ŏI�X�V�t�@�C����
Function latestFile(argDirPath As String)
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    Dim fileTime As Date
    Dim latestTime As Date
    Dim latestFileName As String
    
    Do While currentFileName <> ""
        '�t�H���_���t�@�C���������o��
        fileTime = FileDateTime(argDirPath & "\" & currentFileName) '�擾�����t�@�C���̓������擾
    
        If fileTime > latestTime Then
            '���̔�r�p
            latestTime = fileTime
            '�߂�l�p
            latestFileName = currentFileName
        
        End If
        '���̃t�@�C���������o���i�Ȃ���΃u�����N�j
        currentFileName = Dir()
    Loop
    
    latestFile = latestFileName
End Function
'#�����F�u���z��̓��e���ׂĂ��e�L�X�g�ɓK�p����
'#�����FargReplaceArray:�u���z��i2������1,3��ڂɒu���O�E��Ƃ���j�AargText:�u���O�̕�����
'#�ߒl�F�u���㕶����
Function replaceWithArray(argReplaceArray As Variant, argText As String)

    Dim minIndex As Integer: minIndex = LBound(argReplaceArray, 1)
    Dim replaceText As String: replaceText = argText

    '�ŏ��̗v�f����̏ꍇ�A�����w��͂Ȃ��Ɣ��f���A���������̂܂ܕԂ�
    If argReplaceArray(minIndex, 1) = "" Then
        replaceWithArray = argText
        Exit Function
    End If
    
    '�z��̗v�f������������
    For i = minIndex To UBound(argReplaceArray)
        replaceText = Replace(replaceText, argReplaceArray(i, 1), argReplaceArray(i, 3))
    Next
    replaceWithArray = replaceText
End Function
'#�����F�󂯎�������e�Ńe�L�X�g�t�@�C�����쐬����
'#�����FargFilePath:�t�@�C���̃t���p�X�AargContents:�t�@�C���̓��e�AargCharSet:�����R�[�h
'#�ߒl�F�Ȃ�
Function createTextFile(argFilePath As String, argContents As String, argCharSet As String)
    With CreateObject("ADODB.Stream")
        .charset = argCharSet
        'https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/lineseparatorsenum?view=sql-server-ver16
        .LineSeparator = 10
        .Open
        .WriteText argContents, 0
        If argCharSet = "UTF-8" Then
            'Stream�I�u�W�F�N�g�̐擪����̈ʒu���w�肷��BType�ɒl��ݒ肷��Ƃ���0�ł���K�v������
            .Position = 0
            '�����f�[�^��ނ��o�C�i���f�[�^�ɕύX����
            .Type = 1
            '�ǂݎ��J�n�ʒu�H��3�o�C�g�ڂɈړ�����i3�o�C�g��BOM�t���������폜���邽�߁j
            .Position = 3
            '�o�C�g�������ꎞ�ۑ�
            bytetmp = .read
            '�����ł͕ۑ��͕s�v�B��x���ď������񂾓��e�����Z�b�g����ړI������
            .Close
            '�ēx�J����
            .Open
            '�o�C�g�`���ŏ�������
            .write bytetmp
        End If
        '�ۑ�
        .SaveToFile argFilePath, 2
        '�R�s�[��t�@�C�������
        .Close
    End With
End Function
'#�����F�e�L�X�g�t�@�C�����Ɏw�蕶���񂪂��邩�ǂ������ׂ�B����FTrue�A�Ȃ��FFalse
'#�����FargFilePath:�t�@�C���̃t���p�X�AsearchText:�T������������
'#�ߒl�F�^�U�l�itrue:����Afalse:�Ȃ��j
Function isInTextFile(argFilePath As String, argSearchArray As Variant)

    '�t�@�C�����e�Ǎ�
    Dim fileText As String: fileText = getFileText(argFilePath)
    '����������
    Dim searchValue As String
    '��������������
    Dim findedText As String
    
    For i = LBound(argSearchArray, 1) To UBound(argSearchArray, 1)
        searchValue = argSearchArray(i, 1)
        
        If searchValue <> "" And InStr(fileText, searchValue) > 0 Then
            findedText = findedText & searchValue & ","
        End If
    Next
    isInTextFile = deleteEndText(findedText)
End Function
'#�����F�e�L�X�g�t�@�C�����Ɏw�蕶���񂪂��邩�ǂ������ׂ�B
'#�����FargFilePath:�t�@�C���̃t���p�X�AsearchText:�T������������
'#�ߒl�F�^�U�l�itrue:����Afalse:�Ȃ��j
Function isInExcelFile(argFilePath As String, argSearchArray As Variant) 'searchText As String)
    '�����p
    Dim ws As Worksheet
    '�����͈͍ŏ��Z��
    Dim firstRange As Range
    '�����͈͍ŏI�Z��
    Dim lastRange As Range
    '�u�b�N���J���đS�V�[�g�𑖍�
    Set wb = Workbooks.Open(fileName:=argFilePath, UpdateLinks:=0)
        
    Dim searchValue As String
    Dim findedAddress As String

    For i = 1 To wb.Worksheets.count
        Set ws = wb.Worksheets(i)
        Set firstRange = ws.Range(Split(ws.UsedRange.Address, ":")(0))
        Set lastRange = usedLastRange(ws)

        For j = firstRange.row To lastRange.row
            For k = firstRange.Column To lastRange.Column
                For l = LBound(argSearchArray, 1) To UBound(argSearchArray, 1)
                    searchValue = argSearchArray(l, 1)
                    If searchValue <> "" And InStr(ws.Cells(j, k), searchValue) > 0 Then
                        findedAddress = findedAddress & "[" & ws.Name & ":" & ws.Cells(j, k).Address & ":" & searchValue & "],"
                    End If
                Next
            Next
        Next
    Next
    wb.Close SaveChanges:=False
    isInExcelFile = deleteEndText(findedAddress)
End Function
'#�����F�G�N�Z���t�@�C���ɒu�����{���ۑ�����B
'#�����FargFilePath:�t�@�C���̃t���p�X�AargReplaceArray:�u���z��i2������1,3��ڂɒu���O�E�u����l�Ƃ���j
'#�ߒl�F�u�����ʂ̕�����
Function replaceInExcelFile(argFilePath As String, argReplaceArray As Variant)
    '�����p
    Dim ws As Worksheet
    '�����͈͍ŏ��Z��
    Dim firstRange As Range
    '�����͈͍ŏI�Z��
    Dim lastRange As Range
    '�u�b�N���J���đS�V�[�g�𑖍�
    Set wb = Workbooks.Open(fileName:=argFilePath, UpdateLinks:=0)

    Dim searchValue As String
    Dim replaceValue As String
    Dim replacedAddress As String
    
    For i = 1 To wb.Worksheets.count
        Set ws = wb.Worksheets(i)
        Set firstRange = ws.Range(Split(ws.UsedRange.Address, ":")(0))
        Set lastRange = usedLastRange(ws)

        For j = firstRange.row To lastRange.row
            For k = firstRange.Column To lastRange.Column
                For l = LBound(argReplaceArray, 1) To UBound(argReplaceArray, 1)
                    searchValue = argReplaceArray(l, 1)
                    replaceValue = argReplaceArray(l, 3)

                    If searchValue <> "" And InStr(ws.Cells(j, k), searchValue) > 0 Then
                        '�u������
                        ws.Cells(j, k) = Replace(ws.Cells(j, k), searchValue, replaceValue)
                        replacedAddress = replacedAddress & "[" & ws.Name & ":" & ws.Cells(j, k).Address & _
                                                                    ":" & searchValue & ">" & replaceValue & "],"
                        '�g�����肪������΁A�����ŃJ�X�^�����O�o�͊֐����ĂԂ悤�ɂ���
                    End If
                Next
            Next
        Next
    Next
    wb.Close SaveChanges:=True
    replaceInExcelFile = deleteEndText(replacedAddress)
End Function
'#�����F�t�@�C���p�X�̃e�L�X�g�t�@�C�����쐬���A���̍ۂɒu���������s���B
'#�����FargFilePath:�t�@�C���̃t���p�X�AsearchText:����������AreplaceText:�u��������
'#�ߒl�F�Ȃ�
Function replaceInTextFile(argSrcFilePath As String, argDistFilePath As String, argReplaceArray As Variant)
    '�����R�[�h����i�㏑���ΏۂƓ��������R�[�h�ɂ��邽�߁j
    Dim charset As String: charset = judgeFileCharSet(argSrcFilePath)
    
    Dim searchValue As String
    Dim findedValue As String
    
    If charset = "UTF8" Then
        charset = "UTF-8"
    ElseIf charset = "SJIS" Then
        charset = "SHIFT-JIS"
    End If
    
    With CreateObject("ADODB.Stream")
        .charset = charset
        .Open
        '�R�s�[���t�@�C�����J��
        .LoadFromFile argSrcFilePath
        '�e�L�X�g�`���œ��e���ꎞ�ۑ�
        buf = .ReadText
        
        '������u��
        For l = LBound(argReplaceArray, 1) To UBound(argReplaceArray, 1)
            searchValue = argReplaceArray(l, 1)
            If searchValue <> "" And InStr(buf, searchValue) > 0 Then
                buf = Replace(buf, searchValue, argReplaceArray(l, 3))
                findedValue = findedValue & searchValue & ","
            End If
        Next
        
        With CreateObject("ADODB.Stream")
            .charset = charset
            'https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/lineseparatorsenum?view=sql-server-ver16
            .LineSeparator = 10
            .Open
            .WriteText buf, 0
            If charset = "UTF-8" Then
                'Stream�I�u�W�F�N�g�̐擪����̈ʒu���w�肷��BType�ɒl��ݒ肷��Ƃ���0�ł���K�v������
                .Position = 0
                '�����f�[�^��ނ��o�C�i���f�[�^�ɕύX����
                .Type = 1
                '�ǂݎ��J�n�ʒu�H��3�o�C�g�ڂɈړ�����i3�o�C�g��BOM�t���������폜���邽�߁j
                .Position = 3
                '�o�C�g�������ꎞ�ۑ�
                bytetmp = .read
                '�����ł͕ۑ��͕s�v�B��x���ď������񂾓��e�����Z�b�g����ړI������
                .Close
                '�ēx�J����
                .Open
                '�o�C�g�`���ŏ�������
                .write bytetmp
            End If
            '�ۑ�
            .SaveToFile argDistFilePath, 2
            '�R�s�[��t�@�C�������
            .Close
        End With
        '�R�s�[���t�@�C�������
        .Close
    End With
    replaceInTextFile = deleteEndText(findedValue)
End Function
