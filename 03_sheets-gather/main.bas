Attribute VB_Name = "main"
'###############################
'�@�\���FExcel���W��}�N��
'Author�Fokayasu jun
'�쐬���F2021/10/03
'�X�V���F2023/02/19
'COMMENT�F2023/02����GitHub�ŊǗ����܂��B
'###############################

'�����s����
'�t�@�C�����]�L��A�h���X
Public fileNameCopyAddress As String
'�t�@�C�������V�[�g���ɂ��邩�ݒ�
Public sheetNameFromFileNameFlag As Boolean
'�t�@�C�������u�����N�ɒu�����镶����
Public replaceTextToBlank As String
'�����u�b�N�p�X
Dim newBookPath As String
'�����u�b�N
Dim newBook As Workbook
'�W��V�[�g���z��
Dim copySheetArray As Variant
'�V�[�g�^�C�v�i���O/�ԍ��j
Dim copyType As String
'#���s
Sub exe()
    '���s
    Call main
    '�����ʒm
    MsgBox noticeCount & "�V�[�g���W�񂵂܂����B" & vbCrLf & newBookPath & "�Ɋi�[���Ă��܂��B"
End Sub
'#���C������
Function main()
    '���ʏ����������iModule2�j
    Set customSettingCurrentRange = initialize
    
    '�Ǝ������������iModule1�j
    Call initializeInCustom(customSettingCurrentRange)
    
    '�t�@�C�����ƂɃ`�F�b�N���s
    Call scanLoopWithFile(srcDirPath)
    
    '�W�񌋉ʃt�H���_�̕ۑ�&�N���[�Y
    newBook.SaveAs newBookPath
    newBook.Close SaveChanges:=False
        
    '���s����
    noticeCount = logWriteLine - 2
    
    '�I������
    Call finally
End Function
'#�@�\�Ǝ�����������
Function initializeInCustom(customSettingCurrentRange As Range)

    '�V�[�g���A�V�[�gNo�̏����擾
    Dim sheetNameArrayPerSheet As Variant: sheetNameArrayPerSheet = Split(customSettingCurrentRange.Offset(0, 1).value, ",")
    Dim sheetNoArrayPerSheet As Variant: sheetNoArrayPerSheet = Split(customSettingCurrentRange.Offset(0, 6).value, ",")
    
    '�W��ΏۃV�[�g���
    copySheetArray = IIf(UBound(sheetNameArrayPerSheet) >= 0, sheetNameArrayPerSheet, sheetNoArrayPerSheet)
    copyType = IIf(UBound(sheetNameArrayPerSheet) >= 0, "name", "no")

    '�t�@�C������]�L����A�h���X
    fileNameCopyAddress = getBottomEndRange(customSettingCurrentRange, 1).Offset(0, 1).value
    '�t�@�C�������V�[�g���ɂ��邩�ݒ�
    sheetNameFromFileNameFlag = getBottomEndRange(customSettingCurrentRange, 1).Offset(1, 1).value = "����"
    '�t�@�C�������u�����N�ɒu�����镶����
    replaceTextToBlank = getBottomEndRange(customSettingCurrentRange, 1).Offset(2, 1).value


    '���ʕۑ��p�V�K�u�b�N����
    time = Format(Now(), "yyyy-mm-dd-hh-mm-ss")
    newBookPath = ThisWorkbook.path & "\" & "�W�񌋉�_" & time & ".xlsx"
    Set newBook = Workbooks.Add

    logSheet.Cells.Clear
    logSheet.Cells(1, 1) = "No."
    logSheet.Cells(1, 2) = "�t�H���_"
    logSheet.Cells(1, 3) = "�t�@�C��"
    logSheet.Cells(1, 4) = "���V�[�g��"
    logSheet.Cells(1, 5) = "��V�[�g��"
    logSheet.Cells(1, 6) = "�R�s�[����"
    logSheet.Cells(1, 7) = "����"
End Function
'#�Ώۂ̑S�t�@�C���𑖍�����B�I�v�V�����ɉ����čċA�������s���B
Function scanLoopWithFile(argDirPath As String)
    '�t�H���_���̍ŏ��̃t�@�C�������擾
    Dim currentFileName As String: currentFileName = Dir(argDirPath & "\*.*")
    
    '�����J���Ă���book�̊i�[��
    Dim wb As Workbook
    '�R�s�[���̃V�[�g��
    Dim srcSheetName As String
    '�R�s�[��̃V�[�g��
    Dim distSheetName As String
    '�R�s�[���u�b�N�̊g���q
    Dim srcBookExt As String

    Do While currentFileName <> ""
    
        '�������t�@�C����������ʉ߂��邩�ǂ���
        If isPassFile(argDirPath, currentFileName) Then

            Set wb = Workbooks.Open(fileName:=argDirPath & "\" & currentFileName, UpdateLinks:=0)
            For Each target In copySheetArray

                srcBookExt = objFSO.GetExtensionName(argDirPath & "\" & currentFileName)
            
                '�V�[�g���E�ԍ������ꂩ�ŏ������{�B
                If copyType = "name" And isExistCheckToSheet(wb, target) Then
                    '�V�[�g���w��
                    wb.Sheets(target).Copy After:=newBook.Sheets(newBook.Sheets.count)
                    srcSheetName = wb.Sheets(target).Name
                ElseIf copyType = "no" And isExistCheckToSheet(wb, target) Then
                    '�V�[�g�ԍ��w��
                    wb.Sheets(CInt(target)).Copy After:=newBook.Sheets(newBook.Sheets.count)
                    srcSheetName = wb.Sheets(CInt(target)).Name
                End If

                '�V�[�g�����t�@�C�����ɂ��鏈����ʂ�Ȃ��ꍇ�̂���
                distSheetName = newBook.Sheets(newBook.Sheets.count).Name

                '�t�@�C�������R�s�[��V�[�g�̂ǂ����̃Z���ɓ]�L����ꍇ
                If isExistCheckToSheet(wb, target) And fileNameCopyAddress <> "" Then
                    newBook.Sheets(newBook.Sheets.count).Range(fileNameCopyAddress) = wb.Name
                End If

                '�V�[�g�����t�@�C�����ɂ���ꍇ
                If isExistCheckToSheet(wb, target) And sheetNameFromFileNameFlag Then
                    distSheetName = Replace(wb.Name, replaceTextToBlank, "")
                    '�g���q�폜�i�h�b�g�������j
                    distSheetName = Replace(distSheetName, "." & srcBookExt, "")
                    
                    '�V�[�g���̋֎~�������폜
                    distSheetName = replaceTabooStrWithSheetName(distSheetName)
                    
                    '�v�f����1�iUbound���0�j�̏ꍇ�A1�t�@�C���ɂ�1�V�[�g�ɂȂ�̂Ńt�@�C�����݂̂Ƃ���
                    If UBound(copySheetArray) = 0 Then
                        newBook.Sheets(newBook.Sheets.count).Name = distSheetName
                    Else
                        '�t�@�C���� + �V�[�g���̌`���B��ӂɂ��邽��
                        distSheetName = distSheetName & "-" & srcSheetName
                        newBook.Sheets(newBook.Sheets.count).Name = distSheetName
                    End If
                End If

                '���O�L�^
                If isExistCheckToSheet(wb, target) And logFlag Then
                    logSheet.Cells(logWriteLine, 1) = logWriteLine - 1
                    logSheet.Cells(logWriteLine, 2) = argDirPath & "\"
                    logSheet.Cells(logWriteLine, 3) = currentFileName
                    logSheet.Cells(logWriteLine, 4) = srcSheetName
                    logSheet.Cells(logWriteLine, 5) = distSheetName
                    logSheet.Cells(logWriteLine, 6) = target
                    logSheet.Cells(logWriteLine, 7) = Format(Now(), "yyyy/mm/dd hh:mm:ss")
                    logWriteLine = logWriteLine + 1
                Else
                    '�R�s�[���u�b�N�ɃV�[�g���Ȃ��A�R�s�[���Ă��Ȃ��ꍇ�����O�Ɏc�����ǂ���
                End If

            Next
            wb.Close SaveChanges:=False
            '�}�ȏ��������΍�Ƃ��ĕۑ�
            Debug.Print newBookPath
            newBook.SaveAs newBookPath

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
    
    '�񕝒���
    logSheet.Columns("A:G").AutoFit
    
End Function
