Attribute VB_Name = "main"
'###############################
'�@�\���F�e�L�X�g�t�@�C���u���R�s�[v2
'Author�Fokayasu jun
'�쐬���F2022/04/05
'�X�V���F2023/02/19
'COMMENT�F
'###############################

'�����s����
'�V�[�g�w��
Dim replaceArray As Variant
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
    
    '�t�@�C�����ƂɃ`�F�b�N���s
    Call scanLoopWithFile(srcDirPath)
        
    '���s����
    noticeCount = logWriteLine - 2
    
    '�I������
    Call finally
End Function
'#�@�\�Ǝ�����������
Function initializeInCustom(customSettingCurrentRange As Range)


    '�����s����
    Dim replaceStartRange As Range
    Set replaceStartRange = customSettingCurrentRange.Offset(0, 1)
    Dim replaceEndRange As Range
    Set replaceEndRange = regionEndRange(replaceStartRange, headerFlag:=True, rightTimes:=1)
    replaceArray = IIf(replaceStartRange.value = "", replaceStartRange, topSheet.Range(replaceStartRange, replaceEndRange))
        
    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "���t�H���_"
    logSheet.Cells(1, 3) = "���t�@�C����"
    logSheet.Cells(1, 4) = "��t�H���_"
    logSheet.Cells(1, 5) = "��t�@�C����"
    logSheet.Cells(1, 6) = "�����R�[�h"
    logSheet.Cells(1, 7) = "����"
End Function
'#�Ώۂ̑S�t�@�C���𑖍�����B�I�v�V�����ɉ����čċA�������s���B
Function scanLoopWithFile(argDirPath As String)
    '�t�H���_���̍ŏ��̃t�@�C�������擾
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    Do While currentFileName <> ""
    
        '�������t�@�C����������ʉ߂��邩�ǂ���
        If isPassFile(argDirPath, currentFileName) Then
        
            '�u���t�@�C������
            Call replaceFileCopy(argDirPath, currentFileName)
        
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
    
    '�񕝒���
    logSheet.Columns("A:G").AutoFit
    
End Function
'�t�@�C���ɒu�����������܂��R�s�[����
Function replaceFileCopy(argDirPath As String, currentFileName As String)
    
    '�R�s�[���t�@�C��
    Dim srcFileName As String: srcFileName = currentFileName
    
    '�R�s�[��t�@�C��
    Dim distFileName As String
    
    '�����R�[�h�̔��茋��
    Dim judgedCharSet As String
    
    '�R�s�[�����t���p�X�Ŋi�[����
    Dim srcFilePath As String: srcFilePath = argDirPath & "\" & srcFileName
    
    '�R�s�[����t���p�X�Ŋi�[����
    Dim distFilePath As String
    
    '�u���㕶����i�[�p
    Dim replacedContents As String
    
    '�������t�@�C�����u���Ώۂɂ���Ώ������s��
    For i = LBound(replaceArray) To UBound(replaceArray)
        
        If srcFileName = replaceArray(i, 1) Then
            '�R�s�[��t�@�C����
            distFileName = replaceTabooStrWithFileName(CStr(replaceArray(i, 2)))
            
            '�R�s�[��t�@�C���p�X
            distFilePath = distDirPath & "\" & distFileName
            
            '�R�s�[���̕����R�[�h����
            judgedCharSet = judgeFileCharSet(srcFilePath)
            
            '�u���O������擾
            replacedContents = getFileText(srcFilePath)
            
            '������u���i�L�ڕ����ׂĎ��{����j
            For j = 3 To UBound(replaceArray, 2) Step 2
                replacedContents = Replace(replacedContents, replaceArray(i, j), replaceArray(i, j + 1))
            Next
            
            If judgedCharSet = "UTF8" Then
                Call createTextFile(distFilePath, replacedContents, "UTF-8")
            ElseIf judgedCharSet = "SJIS" Then
                Call createTextFile(distFilePath, replacedContents, "SHIFT-JIS")
            Else
                judgedCharSet = "�����R�[�h�s���ɂ�薢���{"
            End If
            
            '���O�ɋL�^
            logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
            logSheet.Cells(logWriteLine, 2) = argDirPath
            logSheet.Cells(logWriteLine, 3) = srcFileName
            logSheet.Cells(logWriteLine, 4) = distDirPath
            logSheet.Cells(logWriteLine, 5) = distFileName
            logSheet.Cells(logWriteLine, 6) = judgedCharSet
            logSheet.Cells(logWriteLine, 7) = time
            logWriteLine = logWriteLine + 1
        End If
    Next
End Function

