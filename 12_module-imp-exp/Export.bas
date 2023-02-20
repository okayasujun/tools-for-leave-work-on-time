Attribute VB_Name = "Export"
'#�h��/�ӈӁFhttps://vbabeginner.net/bulk-export-of-standard-modules/
'�Q�Ɛݒ�FMicrosoft Visual Basic for Application Extensibilly 5.3�@��ǉ�
'�A�N�e�B�u�u�b�N�̃��W���[�����u�b�N�Ɠ��t�H���_��export����
Sub ExportAll()
    '���W���[��
    Dim module As VBComponent
    '�S���W���[��
    Dim moduleList As VBComponents
    '���W���[���̊g���q
    Dim extension
    '�����Ώۃu�b�N�p�X
    Dim bookPath As String
    '�G�N�X�|�[�g�Ώۃt�@�C���p�X
    Dim exportFilePath  As String
    '�����Ώۃu�b�N
    Dim TargetBook As Workbook
    '���O�����o���s
    Dim logWriteLine As Integer: logWriteLine = 2
    '���[�U�ԓ�
    Dim response As VbMsgBoxResult
    'common���W���[���G�N�X�|�[�g�Ώۃt���O
    Dim commonFlag As Boolean
    
    response = MsgBox("���ʃ��W���[�����G�N�X�|�[�g���܂����H", vbYesNoCancel + vbQuestion)
    
    If response = vbCancel Then
        Exit Sub
    
    ElseIf response = vbNo Then
        commonFlag = False
        
    ElseIf response = vbYes Then
        commonFlag = True
        
    End If
    
    'log�V�[�g�̏�����
    Call logSetUp
    
    If (Workbooks.Count = 1) Then
        '�J���Ă���u�b�N�����u�b�N�݂̂ł���΂����ΏۂƂ���
        Set TargetBook = ThisWorkbook
    Else
        '�����u�b�N���J���Ă���Ώ������s���̃A�N�e�B�u�u�b�N��ΏۂƂ���
        Set TargetBook = ActiveWorkbook
    End If
    
    '�����Ώۃu�b�N�̃p�X���擾
    bookPath = TargetBook.Path
    
    '�����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = TargetBook.VBProject.VBComponents
    
    'VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList
        
        If (module.Type = vbext_ct_ClassModule) Then
            '�N���X
            extension = "cls"
        
        ElseIf (module.Type = vbext_ct_MSForm) Then
            '�t�H�[���@���u.frx�v���ꏏ�ɃG�N�X�|�[�g�����
            extension = "frm"
        
        ElseIf (module.Type = vbext_ct_StdModule) Then
            '�W�����W���[��
            extension = "bas"
        Else
            '���̑� �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE
        End If
        
        If module.Name = "common" And Not commonFlag Then
            '���ʃ��W���[�����G�N�X�|�[�g���Ȃ�
            GoTo CONTINUE
        End If
        
        '�G�N�X�|�[�g���{
        exportFilePath = bookPath & "\" & module.Name & "." & extension
        Call module.Export(exportFilePath)
        
        '�o�͐�m�F�p���O�o��
        Debug.Print exportFilePath
        ThisWorkbook.Worksheets("log").Cells(logWriteLine, 1) = logWriteLine - 1
        ThisWorkbook.Worksheets("log").Cells(logWriteLine, 2) = exportFilePath
        ThisWorkbook.Worksheets("log").Cells(logWriteLine, 3) = "export"
        ThisWorkbook.Worksheets("log").Cells(logWriteLine, 4) = Now()
        logWriteLine = logWriteLine + 1
CONTINUE:
    Next
    
    '�񕝒���
    ThisWorkbook.Worksheets("log").Columns("A:D").AutoFit
End Sub
'���O�V�[�g�̏�����
Function logSetUp()
    ThisWorkbook.Worksheets("log").Cells.Clear
    ThisWorkbook.Worksheets("log").Cells(1, 1) = "No"
    ThisWorkbook.Worksheets("log").Cells(1, 2) = "�t�@�C����"
    ThisWorkbook.Worksheets("log").Cells(1, 3) = "�������"
    ThisWorkbook.Worksheets("log").Cells(1, 4) = "���s����"
End Function
