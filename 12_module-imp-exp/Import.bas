Attribute VB_Name = "Import"
'#�h��/�ӈӁFhttps://vbabeginner.net/bulk-import-of-standard-modules/
'�Q�Ɛݒ�FMicrosoft Visual Basic for Application Extensibilly 5.3�@��ǉ�
'�Q�Ɛݒ�FMicrosoft Scripting Runtime�@��ǉ�
'�A�N�e�B�u�u�b�N�z���ɂ���S���W���[����import����
Sub ImportAll()
    On Error Resume Next
    '�Ώۃu�b�N�i�������̓G�N�X�v���[���[�������H�j
    Dim TargetBook As Workbook
    Set TargetBook = ActiveWorkbook
    '�C���|�[�g�Ώۃt�@�C���p�X
    Dim importDirPath As String
    importDirPath = TargetBook.Path
    '�t�@�C������I�u�W�F�N�g
    Dim objFSO As New FileSystemObject
    '���W���[�����z��
    Dim modulePathArray() As String
    '���W���[���i���[�v�Ŏg�p����s����Variant�^�K�{�j
    'Dim modulePath  As Variant
    '���W���[���g���q
    Dim extension As String
    '���O�����o���s
    Dim logWriteLine As Integer: logWriteLine = 2
    '���[�U�ԓ�
    Dim response As VbMsgBoxResult
    
    response = MsgBox("�����̃��W���[���͏㏑�����܂��B��낵���ł����H", vbOKCancel, "�㏑���m�F")
    If response <> vbOK Then
        Exit Sub
    End If
    
    'log�V�[�g�̏�����
    Call logSetUp
    
    '�z��v�f���w��
    ReDim modulePathArray(0)
    
    '�Ώۃt�H���_�z���̑S���W���[���̃t�@�C���p�X�����W
    Call searchAllFile(importDirPath, modulePathArray)
    
    '�S���W���[���p�X�����[�v
    For Each importFilePath In modulePathArray
        
        '�g���q���������Ŏ擾
        extension = LCase(objFSO.GetExtensionName(importFilePath))
        
        '�g���q��cls�Afrm�Abas�̂����ꂩ�̏ꍇ
        If (extension = "cls" Or extension = "frm" Or extension = "bas") Then
            '�������W���[�����폜
            Call TargetBook.VBProject.VBComponents.Remove(TargetBook.VBProject.VBComponents(objFSO.GetBaseName(importFilePath)))
            '���W���[����ǉ�
            Call TargetBook.VBProject.VBComponents.Import(importFilePath)
        
            '���O�o��
            ThisWorkbook.Worksheets("log").Cells(logWriteLine, 1) = logWriteLine - 1
            ThisWorkbook.Worksheets("log").Cells(logWriteLine, 2) = importFilePath
            ThisWorkbook.Worksheets("log").Cells(logWriteLine, 3) = "import"
            ThisWorkbook.Worksheets("log").Cells(logWriteLine, 4) = Now()
            logWriteLine = logWriteLine + 1
        End If
    Next
    
    '�񕝒���
    ThisWorkbook.Worksheets("log").Columns("A:D").AutoFit
End Sub
'�w��t�H���_�z���̑S�t�@�C���p�X���擾����
'argImportDirPath:�����Ώۃ��[�g�t�H���_�p�X�AargModulePathArray():�t�@�C���p�X�z��
Function searchAllFile(argImportDirPath As String, argModulePathArray() As String)
    '�t�@�C������I�u�W�F�N�g
    Dim objFSO As New FileSystemObject
    '�����Ώۃt�H���_
    Dim dir As Folder
    '�����ΏۃT�u�t�H���_
    Dim subDir As Folder
    '�t�@�C��
    Dim file As file
    '�z��C���f�b�N�X
    Dim i As Integer: i = 0
    
    If Not objFSO.FolderExists(argImportDirPath) Then
        '�t�H���_���Ȃ���ΏI��
        Exit Function
    End If
    
    '�����Ώۃt�H���_�̎擾
    Set dir = objFSO.GetFolder(argImportDirPath)
    
    '�T�u�t�H���_���ċA����
    For Each subDir In dir.SubFolders
        Call searchAllFile(subDir.Path, argModulePathArray)
    Next
    
    '�p�X�z��̗v�f�����擾
    i = UBound(argModulePathArray)
    
    '�������t�H���_���̃t�@�C�����擾
    For Each file In dir.Files
    
        '�v�f�����łɂ��邩�ǂ����B�����True��Ԃ�
        If (i <> 0 Or argModulePathArray(i) <> "") Then
            i = i + 1
            '�v�f�l��ێ������܂ܗv�f���𑝉�
            ReDim Preserve argModulePathArray(i)
        End If
        
        '�t�@�C���p�X��z��Ɋi�[�i�����ł͊g���q�����肵�Ȃ��j
        argModulePathArray(i) = file.Path
    Next
End Function
'���O�V�[�g�̏�����
Function logSetUp()
    ThisWorkbook.Worksheets("log").Cells.Clear
    ThisWorkbook.Worksheets("log").Cells(1, 1) = "No"
    ThisWorkbook.Worksheets("log").Cells(1, 2) = "�t�@�C����"
    ThisWorkbook.Worksheets("log").Cells(1, 3) = "�������"
    ThisWorkbook.Worksheets("log").Cells(1, 4) = "���s����"
End Function
