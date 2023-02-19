Attribute VB_Name = "main"
'###############################
'�@�\���F�t�@�C�����l�[���}�N��
'Author�Fokayasu jun
'�쐬���F2022/04/17
'�X�V���F2023/02/19
'COMMENT�F
'###############################

'�����s�P�[�X�P
'�A�ԕt�^
Public serialNoFlag As Boolean
'�ړ���
Public prefix As String
'�ڔ���
Public suffix As String
'�u������
Public replaceConditionArray As Variant
'�����s�P�[�X�Q
'�u�����X�g
Public renameListArray As Variant
'���X�g�̊J�n�Z��
Public renameFileStartRange As Range
'���l�[���O�t�@�C�����������ݑΏۃZ��
Public renameWriteRange As Range
'�ԍ��t�^�p
Public serialNo As Integer
'#�`�F�b�N����
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
        Case "replace"
            '���s�P�[�X�P�̃t�@�C�����u������
            Call scanLoopWithFile(srcDirPath, "replace")
            logSheet.Select
        Case "get"
            '���s�P�[�X�Q�̃t�@�C�����擾����
            Call scanLoopWithFile(srcDirPath, "get")
        Case "serialNo"
            '���s�P�[�X�Q�̘A�ԕt�^����
            Call setSerialNo
        Case "rename"
            '���s�P�[�X�Q�̃��l�[������
            Call renameFile
            logSheet.Select
    End Select
        
    '���s����
    noticeCount = logWriteLine - 2
    
    '�I������
    Call finally
End Function
'#�@�\�Ǝ�����������
Function initializeInCustom(customSettingCurrentRange As Range)
    
    '�����s�P�[�X�P�̃p�����[�^
    serialNoFlag = customSettingCurrentRange.Offset(0, 1) = "����"
    prefix = customSettingCurrentRange.Offset(1, 1)
    suffix = customSettingCurrentRange.Offset(2, 1)
    replaceConditionArray = Range(customSettingCurrentRange.Offset(3, 1), regionEndRange(customSettingCurrentRange.Offset(3, 1), False, 1))
    
    Set renameFileStartRange = getBottomEndRange(customSettingCurrentRange, 2)
    
    '�����s�P�[�X�Q�̃p�����[�^
    renameListArray = Range(renameFileStartRange.Offset(0, 1), regionEndRange(renameFileStartRange, False, 3))
    Set renameWriteRange = renameFileStartRange
    
    serialNo = 1
    
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "�t�H���_"
    logSheet.Cells(1, 3) = "���l�[���O"
    logSheet.Cells(1, 4) = "���l�[����"
    logSheet.Cells(1, 5) = "������"
    logSheet.Cells(1, 6) = "�t�@�C���̍X�V����"
    logSheet.Cells(1, 7) = "��������"
End Function
'#�Ώۂ̑S�t�@�C���𑖍�����B�I�v�V�����ɉ����čċA�������s���B
Function scanLoopWithFile(argDirPath As String, processType As String)
    '�t�H���_���̍ŏ��̃t�@�C�������擾
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '�������t�@�C����������ʉ߂��邩�ǂ���
        If isPassFile(argDirPath, currentFileName) Then
            If processType = "replace" Then
                '�t�@�C�����������Ƃ����R�s�[ or �ړ�
                Call replaceFileName(argDirPath, currentFileName)
                
            ElseIf processType = "get" Then
                '���݂̃t�@�C���������o������
                Call writeFileName(argDirPath, currentFileName) '������
            End If
        End If
        
        '���̃t�@�C���������o���i�Ȃ���΃u�����N�j
        currentFileName = Dir()
    Loop
    
    If recursiveFlag Then
        '�t�H���_���̃T�u�t�H���_�����Ɏ擾
        For Each directory In objFSO.getfolder(argDirPath).SubFolders
            '�ċA����
            Call scanLoopWithFile(directory.Path, processType)
        Next
    End If
End Function
'���s�P�[�X�P�̎g�p
'�ړ����t�H���_�͍ċA�����̏ꍇ�A�O���[�o���ϐ��̒l�ƈقȂ�\�������邽�ߓK�X��������擾����
Function replaceFileName(argDirPath As String, currentFileName As String)

    '�u�����N�u���㕶��
    Dim afterFileName As String
    '�ړ�/�R�s�[���A��A�`�F�b�N�p�t�@�C�����i�[�p
    Dim srcFileName As String: srcFileName = currentFileName
    Dim distFileName As String ': distFileName = currentFileName
    Dim srcFilePath As String: srcFilePath = argDirPath & "\" & currentFileName
    Dim distFilePath As String
    Dim no As String
    
    '�t�@�C���̊g���q�i�ڔ����t�^�̂��߁j
    Dim currentFileExt As String: currentFileExt = objFSO.GetExtensionName(srcFilePath)
    
    '�u�����{
    distFileName = replaceWithArray(replaceConditionArray, srcFileName)
    '�ړ����t�^
    distFileName = prefix & distFileName
    '�ڔ����t�^
    distFileName = Replace(distFileName, "." & currentFileExt, suffix & "." & currentFileExt)
    '�t�@�C���֎~�����폜
    distFileName = replaceTabooStrWithFileName(distFileName)
    
    '�A��
    If serialNoFlag Then
        distFileName = Format(serialNo, "00") & "_" & distFileName
        serialNo = serialNo + 1
    End If
    
    '���l�[��
    Name argDirPath & "\" & srcFileName As argDirPath & "\" & distFileName
    '���O�p
    distFilePath = argDirPath & "\" & distFileName
    
    If logFlag Then
        '���O�L�^
        logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
        logSheet.Cells(logWriteLine, 2) = argDirPath & "\"
        logSheet.Cells(logWriteLine, 3) = srcFileName
        logSheet.Cells(logWriteLine, 4) = distFileName
        logSheet.Cells(logWriteLine, 5) = "=NOT(EXACT(C" & logWriteLine & ",D" & logWriteLine & "))"
        logSheet.Cells(logWriteLine, 6) = objFSO.getFile(distFilePath).DateLastModified
        logSheet.Cells(logWriteLine, 7) = Format(Now(), "yyyy/mm/dd hh:mm:ss")
        logWriteLine = logWriteLine + 1
    End If
End Function
'�n���ꂽ�t�@�C����������̃Z���ɏ����o���B�Z���Ԓn�̓C���N�������g����i���s�P�[�X�Q�̎g�p�z��j
'�������͎g�p���Ă��Ȃ����A����̉��C�̂��ߎc���Ă���
Function writeFileName(argDirPath As String, currentFileName As String)
    renameWriteRange.Offset(0, 1) = argDirPath
    renameWriteRange.Offset(0, 2) = currentFileName
    Set renameWriteRange = renameWriteRange.Offset(1, 0)
    logWriteLine = logWriteLine + 1
End Function
'�A�Ԃ�t�^����i���s�P�[�X�Q�̎g�p�z��j
Function setSerialNo()
    'renameListArray�͌������擾���邽�߂����i������Ƃ������ȁj�B
    For i = LBound(renameListArray) To UBound(renameListArray)
        renameWriteRange.Offset(0, 6) = Format(serialNo, "00") & "_" & renameWriteRange.Offset(0, 6)
        '�C���N�������g
        Set renameWriteRange = renameWriteRange.Offset(1, 0)
        serialNo = serialNo + 1
        logWriteLine = logWriteLine + 1
    Next
End Function
'���͓��e�ɏ]�����l�[������i���s�P�[�X�Q�̎g�p�z��j
Function renameFile()
    '���l�[���O
    Dim srcDirName As String
    Dim srcFileName As String
    '���l�[����
    Dim distDirName As String
    Dim distFileName As String
    
    For i = LBound(renameListArray) To UBound(renameListArray)
        srcDirName = renameListArray(i, 1)
        srcFileName = renameListArray(i, 2)
        distDirName = renameListArray(i, 5)
        distFileName = renameListArray(i, 6)
        
        '�t�@�C���֎~�����폜
        distFileName = replaceTabooStrWithFileName(distFileName)
        Name srcDirName & "\" & srcFileName As distDirName & "\" & distFileName

        If logFlag Then
            '���O�L�^
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = srcDirName & "\"
            logSheet.Cells(logWriteLine, 3) = srcFileName
            logSheet.Cells(logWriteLine, 4) = distFileName
            logSheet.Cells(logWriteLine, 5) = "=NOT(EXACT(C" & logWriteLine & ",D" & logWriteLine & "))"
            logSheet.Cells(logWriteLine, 6) = objFSO.getFile(distDirName & "\" & distFileName).DateLastModified
            logSheet.Cells(logWriteLine, 7) = Format(Now(), "yyyy/mm/dd hh:mm:ss")
            logWriteLine = logWriteLine + 1
        End If
    Next
    
    '�񕝒���
    logSheet.Columns("A:G").AutoFit
    
End Function
