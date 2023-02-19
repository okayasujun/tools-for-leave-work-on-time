Attribute VB_Name = "main"
'###############################
''�@�\���F�t�@�C���d�����}�N��
''Author�Fokayasu jun
''�쐬���F2021/10/30
''�X�V���F2022/12/22
'COMMENT�F
'###############################

'�����s����
'�������
Public processType As String
'�t�@�C�������u�����N�ɒu�����镶����
Public replaceTextToBlank As String
'#�`�F�b�N����
Sub exe()
    '���s
    Call main
    '�����ʒm
    MsgBox noticeCount & "���擾���܂����B"
End Sub
'#���C������
Function main()
    '���ʏ���������
    Set customSettingCurrentRange = initialize
    
    '�Ǝ�����������
    Call initializeInCustom(customSettingCurrentRange)
    
    '�t�@�C�����ƂɃ`�F�b�N���s
    Call scanLoopWithFile(srcDirPath)
        
    '���s����
    noticeCount = logWriteLine - 2
    
    '�I������
    Call finally
End Function
'#�@�\�Ǝ�����������
Function initializeInCustom(customSettingCurrentRange As Range)
    
    processType = customSettingCurrentRange.Offset(0, 1).value
    replaceTextToBlank = customSettingCurrentRange.Offset(1, 1).value
    
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "���t�H���_"
    logSheet.Cells(1, 3) = "���t�@�C����"
    logSheet.Cells(1, 4) = "��t�H���_"
    logSheet.Cells(1, 5) = "��t�@�C����"
    logSheet.Cells(1, 6) = "�������"
    logSheet.Cells(1, 7) = "����"
End Function
'#�Ώۂ̑S�t�@�C���𑖍�����B�I�v�V�����ɉ����čċA�������s���B
Function scanLoopWithFile(argDirPath As String)
    '�t�H���_���̍ŏ��̃t�@�C�������擾
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '�������t�@�C����������ʉ߂��邩�ǂ���
        If isPassFile(argDirPath, currentFileName) Then
            '�t�@�C�����������Ƃ����R�s�[ or �ړ�
            Call fileMoveOrCopy(argDirPath, currentFileName)
        End If
        
        '���̃t�@�C���������o���i�Ȃ���΃u�����N�j
        currentFileName = Dir()
    Loop
    
    If recursiveFlag Then
        '�t�H���_���̃T�u�t�H���_�����Ɏ擾
        For Each directory In objFSO.getfolder(argDirPath).SubFolders
            '�ċA����
            Call scanLoopWithFile(directory.path)
        Next
    End If
End Function
'�ړ����t�H���_�͍ċA�����̏ꍇ�A�O���[�o���ϐ��̒l�ƈقȂ�\�������邽��
'�K�X��������擾����
Function fileMoveOrCopy(argDirPath As String, currentFileName As String)

    '�ړ�/�R�s�[���A��A�`�F�b�N�p�t�@�C�����i�[�p
    Dim srcFileName As String: srcFileName = currentFileName
    Dim distFileName As String

    '�w�蕶�����u�����N�֒u��
    distFileName = Replace(currentFileName, replaceTextToBlank, "")
    
    If processType = "�ړ�" Then
        '�ړ����@As �ړ���
        Name argDirPath & "\" & srcFileName As distDirPath & "\" & distFileName
    ElseIf processType = "�R�s�[" Then
        '�R�s�[��, �ړ���
        objFSO.CopyFile argDirPath & "\" & srcFileName, distDirPath & "\" & distFileName
    End If
    
    If logFlag Then
        '���O�L�^
        logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
        logSheet.Cells(logWriteLine, 2) = argDirPath & "\"
        logSheet.Cells(logWriteLine, 3) = srcFileName
        logSheet.Cells(logWriteLine, 4) = distDirPath
        logSheet.Cells(logWriteLine, 5) = distFileName
        logSheet.Cells(logWriteLine, 6) = processType
        logSheet.Cells(logWriteLine, 7) = Format(Now(), "yyyy/mm/dd hh:mm:ss")
        logWriteLine = logWriteLine + 1
    End If

End Function
