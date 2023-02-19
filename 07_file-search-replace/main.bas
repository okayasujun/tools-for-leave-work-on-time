Attribute VB_Name = "main"
'###############################
'�@�\���F�����E�u��v1.5
'Author�Fokayasu jun
'�쐬���F2022/12/05
'�X�V���F2022/12/25
'COMMENT�F
'###############################

'�����s����
'���s�P�[�X�P�����i�[�z��
Dim exe1Array As Variant
'���s�P�[�X�Q�����i�[�z��
Dim exe2Array As Variant
'#����
Sub exe()
    '���s
    Call main
    '�����ʒm
    MsgBox noticeCount & "���������܂����B"
End Sub
'#���C������
Function main()
    '���ʏ����������iModule2�j
    Set customSettingCurrentRange = initialize
    
    '�Ǝ������������iModule1�j
    Call initializeInCustom(customSettingCurrentRange)
    
    Select Case Application.Caller
        Case "search1"
            Call scanLoopWithFile(srcDirPath)
        Case "replace1"
            Call scanLoopWithFile(srcDirPath)
        Case "get"
            Call scanLoopWithFile(srcDirPath)
        Case "search2"
            Call search2exe
        Case "replace2"
            Call replace2exe
    End Select
    
    '���s����
    noticeCount = logWriteLine - 2
    
    '�񕝒���
    logSheet.Columns("A:G").AutoFit
    
    '�I������
    Call finally
End Function
'#�@�\�Ǝ�����������
Function initializeInCustom(customSettingCurrentRange As Range)
    '���s�P�[�X�P
    Dim exe1StartRange As Range
    Set exe1StartRange = customSettingCurrentRange.Offset(0, 1)
    exe1Array = IIf(IsEmpty(exe1StartRange.value), _
                    exe1StartRange, _
                    topSheet.Range(exe1StartRange, regionEndRange(exe1StartRange.Offset(0, 1))))

    '���s�P�[�X�Q
    Dim exe2StartRange As Range
    Set exe2StartRange = getBottomEndRange(customSettingCurrentRange, 2).Offset(0, 1)
    exe2Array = IIf(IsEmpty(exe2StartRange.value), _
                        exe2StartRange, _
                        topSheet.Range(exe2StartRange, regionEndRange(exe2StartRange.Offset(0, 1), headerFlag:=True)))


    '�t�@�C���擾�p
    If Application.Caller = "get" Then
        '�t�@�C���������o���ɔ�����
        Set customSettingCurrentRange = exe2StartRange
        '�����l�̃N���A�i�l�̂݁j
        topSheet.Range(customSettingCurrentRange, topSheet.Cells(getBottomEndRange(customSettingCurrentRange, 1).row, 5)).ClearContents
    End If

    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "�t�H���_"
    logSheet.Cells(1, 3) = "�t�@�C����"
    logSheet.Cells(1, 4) = "���o�E�u�����"
    logSheet.Cells(1, 5) = "�����R�[�h"
    logSheet.Cells(1, 6) = "���s�_�@"
    logSheet.Cells(1, 7) = "����"
End Function
'#�Ώۂ̑S�t�@�C���𑖍�����B�I�v�V�����ɉ����čċA�������s���B
Function scanLoopWithFile(argDirPath As String)
    '�t�H���_���̍ŏ��̃t�@�C�������擾
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '�������t�@�C����������ʉ߂��邩�ǂ���
        If isPassFile(argDirPath, currentFileName) Then
            Select Case Application.Caller
                Case "search1"
                    '�ėp��������
                    Call search1exe(argDirPath, currentFileName)
                Case "replace1"
                    '�ėp�u������
                    Call replace1exe(argDirPath, currentFileName)
                Case "get"
                    '�t�@�C�����擾
                    Call writeFileName(argDirPath, currentFileName)
            End Select
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
'#���s�P�[�X�P�̌�������
Function search1exe(argDirPath As String, argFileName As String)
    '�����Ώۂ̔z��͂����œn���i�ʌ�����n���P�[�X�����邽�߁j
    Call mainSearch(argDirPath, argFileName, exe1Array)

End Function
'�������������{����
Function mainSearch(argDirPath As String, argFileName As String, searchArray As Variant)
    '�t�@�C���̃t���p�X
    Dim filePath As String: filePath = argDirPath & "\" & argFileName

    Dim findedAddress As String

    If InStr(objFSO.getFile(filePath).Type, "Excel") > 0 Then
        'Excel�n�t�@�C��
        '�u�b�N�����猟������
        findedAddress = isInExcelFile(filePath, searchArray)
        If findedAddress <> "" Then
            '���O�o��
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = argFileName
            logSheet.Cells(logWriteLine, 4) = findedAddress
            logSheet.Cells(logWriteLine, 5) = "-"
            logSheet.Cells(logWriteLine, 6) = Application.Caller
            time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    Else
        '�e�L�X�g�n�t�@�C����z��
        findedAddress = isInTextFile(filePath, searchArray)
        If findedAddress <> "" Then
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = argFileName
            logSheet.Cells(logWriteLine, 4) = findedAddress
            logSheet.Cells(logWriteLine, 5) = judgeFileCharSet(argDirPath & "\" & argFileName)
            logSheet.Cells(logWriteLine, 6) = Application.Caller
            time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    End If
End Function
'#���s�P�[�X�P�̒u������
Function replace1exe(argDirPath As String, argFileName As String)

    '�u���Ώۂ̔z��͂����œn���i�ʌ�����n���P�[�X�����邽�߁j
    Call mainReplace(argDirPath, argFileName, exe1Array)

End Function
'�u�����������{����
Function mainReplace(argDirPath As String, argFileName As String, replaceArray As Variant)
    '�t�@�C���̃t���p�X
    Dim filePath As String: filePath = argDirPath & "\" & argFileName
    
    Dim findedAddress As String
    
    If InStr(objFSO.getFile(filePath).Type, "Excel") > 0 Then
        'Excel�n�t�@�C��
        '�u�b�N�����猟������
        findedAddress = replaceInExcelFile(filePath, replaceArray)
        If findedAddress <> "" Then
            '���O�o��
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = argFileName
            logSheet.Cells(logWriteLine, 4) = findedAddress
            logSheet.Cells(logWriteLine, 5) = "-"
            logSheet.Cells(logWriteLine, 6) = Application.Caller
            time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    Else
        '�e�L�X�g�n�t�@�C����z��
        findedAddress = replaceInTextFile(filePath, filePath, replaceArray)
        If findedAddress <> "" Then
            '���O�o��
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = argFileName
            logSheet.Cells(logWriteLine, 4) = findedAddress
            logSheet.Cells(logWriteLine, 5) = judgeFileCharSet(argDirPath & "\" & argFileName)
            logSheet.Cells(logWriteLine, 6) = Application.Caller
            time = Format(Now(), "yyyy/mm/dd/ hh:mm:ss")
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    End If
End Function
'#�n���ꂽ�t�@�C����������̃Z���ɏ����o���B�Z���̓C���N�������g����
Function writeFileName(argDirPath As String, currentFileName As String)
    customSettingCurrentRange = argDirPath
    customSettingCurrentRange.Offset(0, 1) = currentFileName
    Set customSettingCurrentRange = customSettingCurrentRange.Offset(1, 0)
    logWriteLine = logWriteLine + 1
End Function
'#���s�P�[�X�Q�̌�������
Function search2exe()
    Dim dirPath As String
    Dim fileName As String
    
    Dim searchArray As Variant
    
    For i = LBound(exe2Array) To UBound(exe2Array)
        dirPath = exe2Array(i, 1)
        fileName = exe2Array(i, 2)
        
        '���s�P�[�X�Q�̌����z������s�P�[�X�P�̂悤�ɕϊ�����
        searchArray = convertArrayFrom2dTo2d(exe2Array, CInt(i), 3)
        
        If dirPath <> "" And fileName <> "" Then
            Call mainSearch(dirPath, fileName, searchArray)
        End If
    Next
End Function
'#���s�P�[�X�Q�̒u������
Function replace2exe()
    Dim dirPath As String
    Dim fileName As String

    Dim replaceArray As Variant
    
    For i = LBound(exe2Array) To UBound(exe2Array)
        dirPath = exe2Array(i, 1)
        fileName = exe2Array(i, 2)
        
        '���s�P�[�X�Q�̒u���z������s�P�[�X�P�̂悤�ɕϊ�����
        replaceArray = convertArrayFrom2dTo2d(exe2Array, CInt(i), 3)

        If dirPath <> "" And fileName <> "" Then
            Call mainReplace(dirPath, fileName, replaceArray)
        End If
        
    Next
End Function
'#�����F�񎟌��z������s�̓����œ񎟌��z��ɕϊ�����
'#�����FsrcArray:�z��AargReadRow:�񎟌��̓���s�AargReadStartColumn:�񎟌��ڂ̊J�n�C���f�b�N�X
'#�ߒl�F�ꎟ���z��
Function convertArrayFrom2dTo2d(srcArray As Variant, argReadRow As Integer, argReadStartColumn As Integer)
    Dim returnArray() As Variant
    '�v�f���̐ݒ�
    ReDim returnArray(UBound(srcArray, 2), 3)
    Dim count As Integer: count = 0
    
    '�񎟌��ڂɂ��ă��[�v���s��
    For i = argReadStartColumn To UBound(srcArray, 2) Step 2

        If srcArray(argReadRow, i) <> "" Then
            '�����l
            returnArray(count, 1) = srcArray(argReadRow, i)
            '�u���l
            returnArray(count, 3) = srcArray(argReadRow, i + 1)
            count = count + 1
        End If
    Next
    convertArrayFrom2dTo2d = returnArray
End Function
