Attribute VB_Name = "A_�t�H���_�쐬"
'�I�u�W�F�N�g���V�[�g�̃V�[�g���iIndex�Œ�`���������������H�j
Public Const OBJECT_SHEET = "�I�u�W�F�N�g"
'�I�u�W�F�N�gmeta�t�@�C���̏����Ǘ�����V�[�g�̃V�[�g��
Public Const OBJECT_META_SHEET = "CustomObject"
'���ڏ��V�[�g�̃V�[�g���iIndex�Œ�`���������������H�j
Public Const ITEM_SHEET = "����"
'���ڂ̃^�O�����Ǘ�����V�[�g�̃V�[�g��
Public Const ITEM_META_SHEET = "CustomItem"
'�y�[�W���C�A�E�g���V�[�g�̃V�[�g��
Public Const LAYOUT_SHEET = "�y�[�W���C�A�E�g"
'�t�H���_�쐬
Sub A_�t�H���_�쐬()
    
    'TODO:�����ƌ����ȃG���[�`�F�b�N���K�v��
    If Not Sheets(OBJECT_SHEET).Cells(4, 4) Like "*__c" Then
        MsgBox "�I�u�W�F�N�g�����u__c�v�ŏI����Ă��܂���B"
        Exit Sub
    End If
    
    '�I�u�W�F�N�g�̐e�t�H���_�쐬
    If Dir(ThisWorkbook.Path & "\objects\", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\objects\"
    End If
    
    '�I�u�W�F�N�g���̃t�H���_�����[�g�Ƃ���
    Dim rootDirName As String: rootDirName = ThisWorkbook.Path & "\objects\" & Sheets(OBJECT_SHEET).Cells(4, 4) & "\"
    
    '�I�u�W�F�N�g���t�H���_
    If Dir(rootDirName, vbDirectory) = "" Then
        MkDir rootDirName
    End If
    
    '�R���p�N�g���C�A�E�g
    If Dir(rootDirName & "compactLayouts\", vbDirectory) = "" Then
        MkDir rootDirName & "compactLayouts\"
    End If
    
    '����
    If Dir(rootDirName & "fields\", vbDirectory) = "" Then
        MkDir rootDirName & "fields\"
    End If
    
    '���X�g�r���[
    If Dir(rootDirName & "listViews\", vbDirectory) = "" Then
        MkDir rootDirName & "listViews\"
    End If
    
    '���͋K��
    If Dir(rootDirName & "validationRules\", vbDirectory) = "" Then
        MkDir rootDirName & "validationRules\"
    End If
    
    '���R�[�h�^�C�v
    If Dir(rootDirName & "recordTypes\", vbDirectory) = "" Then
        MkDir rootDirName & "recordTypes\"
    End If
    
    '�p�X�i����̓I�u�W�F�N�g�z���ł͂Ȃ��j
    If Dir(ThisWorkbook.Path & "\tabs\", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\tabs\"
    End If

'    '�����t�^�̓f�[�^���[�_���炵���������S�Ȃ��ߔ񐄏�
'    If Dir(ThisWorkbook.Path & "\profiles\", vbDirectory) = "" Then
'        MkDir ThisWorkbook.Path & "\profiles\"
'    End If

'    '���C�A�E�g�͉�ʂ���쐬�������������ݒ肪��������邽�ߔ񐄏�
'    If Dir(ThisWorkbook.Path & "\layouts\", vbDirectory) = "" Then
'        MkDir ThisWorkbook.Path & "\layouts\"
'    End If
    
    MsgBox "�t�H���_���쐬���܂����B"
End Sub
