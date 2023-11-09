Attribute VB_Name = "A_����"
'�I�u�W�F�N�g���V�[�g�̃V�[�g���iIndex�Œ�`���������������H�j
Public Const OBJECT_SHEET = "�I�u�W�F�N�g"

'�I�u�W�F�N�gmeta�t�@�C���̏����Ǘ�����V�[�g�̃V�[�g��
Public Const OBJECT_META_SHEET = "CustomObject"

'���ڏ��V�[�g�̃V�[�g���iIndex�Œ�`���������������H�j
Public Const ITEM_SHEET = "����"

'���ڂ̃^�O�����Ǘ�����V�[�g�̃V�[�g��
Public Const ITEM_META_SHEET = "CustomItem"

'�������V�[�g�̃V�[�g��
Public Const PERMISSION_SHEET = "����"

'�y�[�W���C�A�E�g���V�[�g�̃V�[�g��
Public Const LAYOUT_SHEET = "�y�[�W���C�A�E�g"

'True����������
Public Const ON_TRUE = "�Z"

'�e�L�X�g�t�@�C���o�͗p
Public stream As Object

'���K�\��
Public regexp As Object

'�I�u�W�F�N�g���V�[�g
Public objSheet As Worksheet

'�I�u�W�F�N�g���^���V�[�g
Public objMetaSheet As Worksheet

'���ڏ��V�[�g
Public itemSheet As Worksheet

'���ڃ��^���V�[�g
Public itemMetaSheet As Worksheet

'�������V�[�g
Public permissionSheet As Worksheet

'�I�u�W�F�N�gAPI��
Public objApiName As String

'�t�@�C����
Public fileName As String

'���ڃt�H���_�p�X
Public fieldsDirPath As String

'������
Public Function initiarize()
    Set regexp = CreateObject("VBScript.RegExp")
    Set objSheet = Sheets(OBJECT_SHEET)
    Set objMetaSheet = Sheets(OBJECT_META_SHEET)
    Set itemSheet = Sheets(ITEM_SHEET)
    Set itemMetaSheet = Sheets(ITEM_META_SHEET)
    Set permissionSheet = Sheets(PERMISSION_SHEET)
    objApiName = Sheets(OBJECT_SHEET).Cells(4, 4).Value
    fieldsDirPath = ThisWorkbook.path & "\objects\" & objApiName & "\fields\"
End Function
'�e�L�X�g�o�͗p
Public Function openStream()
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "UTF-8"
    stream.Open
End Function
'���K�\��
Public Function setupRegexp(argPattern As String)
    '�u���������o�p�p�^�[���iVBA�ōm���ǂ݂͎g���Ȃ��j
    regexp.Pattern = argPattern
    '�p�啶������������ʂ��Ȃ�
    regexp.IgnoreCase = True
    '������S�̂ɑ΂��ăp�^�[���}�b�`������
    regexp.Global = True
End Function
'UTF-8�ŕۑ�����Ƃ��̕ۑ��������X�g���[��object�ƃt�@�C�����ōs��
Public Function saveTextWithUTF8(stream As Object, fileFullName As String)
    'Stream�I�u�W�F�N�g�̐擪����̈ʒu���w�肷��BType�ɒl��ݒ肷��Ƃ���0�ł���K�v������
    stream.Position = 0
    '�����f�[�^��ނ��o�C�i���f�[�^�ɕύX����
    stream.Type = 1
    '�ǂݎ��J�n�ʒu�H��3�o�C�g�ڂɈړ�����i3�o�C�g��BOM�t���������폜���邽�߁j
    stream.Position = 3
    '�o�C�g�������ꎞ�ۑ�
    bytetmp = stream.Read
    '�����ł͕ۑ��͕s�v�B��x���ď������񂾓��e�����Z�b�g����ړI������
    stream.Close
    '�ēx�J����
    stream.Open
    '�o�C�g�`���ŏ������ނ��
    stream.write bytetmp
    Call checkExistDir(getDirPath(fileFullName))
    '�ۑ�
    stream.SaveToFile fileFullName, 2
    '�R�s�[��t�@�C�������
    stream.Close
End Function
'�t�@�C���p�X�����݂��邩�`�F�b�N����B�Ȃ���΂���
Public Function checkExistDir(path As String)
    '�t�@�C������I�u�W�F�N�g
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    '�t�H���_���z��
    Dim dirs() As String: dirs = Split(path, "\")
    '�����t�H���_
    Dim incrementDir As String
    
    For i = LBound(dirs) To UBound(dirs)
        If incrementDir <> "" And Not objFSO.FolderExists(incrementDir) Then
            objFSO.CreateFolder (incrementDir)
        End If
        incrementDir = incrementDir & dirs(i) & "\"
    Next
End Function
'�t�@�C���̃t���p�X����t�H���_�p�X���擾����
Public Function getDirPath(argFilePath As String)
    Dim dirs As Variant: dirs = Split(argFilePath, "\")
    getDirPath = Left(argFilePath, Len(argFilePath) - Len(dirs(UBound(dirs))) - 1) & "\"
End Function
