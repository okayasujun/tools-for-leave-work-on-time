Attribute VB_Name = "main"
'�c�[�����Q�Ɛݒ肩��uMicrosoft XML, v6.0�v��ON�ɂ���
Dim ws As Worksheet
'�h�L�������g�I�u�W�F�N�g
Dim doc As MSXML2.DOMDocument60
'��������XML�v�f�I�u�W�F�N�g���i�[����
Dim XMLchild As Object '���[�v���̏����Ώۂɂ���āA����I�u�W�F�N�g���قȂ邩��A�ڍׂȃI�u�W�F�N�g�͎w�肵�Ȃ��ł����B
'�����o���s
Dim writeLine As Integer
'�������v�f�̊K�w���x��
Dim level As Integer
'�����L�^�J�n��i�������ɑ������ɉ����ăC���N�������g����j
Dim attributesWriteCol As Integer
Const FILE_PATH = "C:\Users\tp04372\Documents\macro\������\sample.xml"
Sub trial2()
    Set ws = ActiveSheet
    ws.Cells.Clear
    ws.Cells(1, 1) = "�Z��m�[�h��"
    ws.Cells(1, 2) = "�q�m�[�h��"
    ws.Cells(1, 3) = "�e�v�f��"
    ws.Cells(1, 4) = "���x��"
    ws.Cells(1, 5) = "�v�f��"
    ws.Cells(1, 6) = "�v�f���e"
    ws.Cells(1, 7) = "�v�f�^�C�v"
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    '�Ώۃt�@�C���ǂݍ���
    doc.Load FILE_PATH
    '���ꂼ�ꏉ����
    writeLine = 1
    level = 0
    attributesWriteCol = 7
    
    Call writeXMLElement(doc)
    '�������B
    Set doc = Nothing
End Sub
'#�n���ꂽXML�e�v�f�ɂ��čċA�I�Ɏq�v�f���Ăяo�����̓��e���Z���ɏ����o��
Private Sub writeXMLElement(XMLparent As Variant)
'    Debug.Print XMLparent.ChildNodes.Length
    
    'Debug.Print XMLparent.parseError.reason
    
    If XMLparent.ChildNodes.Length <> 0 Then
        For Each XMLchild In XMLparent.ChildNodes
        
'            Debug.Print XMLparent.ChildNodes.Length
'            Debug.Print XMLchild.ChildNodes.Length
            '�v�f���e�̏ꍇ�A�q�m�[�h����0�ɂȂ�B�o�͂̈Ӗ����Ȃ�����X�L�b�v������
            If XMLparent.ChildNodes.Length <> 0 Then
                level = level + 1
                writeLine = writeLine + 1
                ws.Cells(writeLine, 1) = XMLparent.ChildNodes.Length '����̈Ӗ������܂����킩���Ă��Ȃ�
                ws.Cells(writeLine, 2) = XMLchild.ChildNodes.Length
                ws.Cells(writeLine, 3) = XMLchild.BaseName
                ws.Cells(writeLine, 4) = level
                ws.Cells(writeLine, 5) = XMLchild.nodeName
                ws.Cells(writeLine, 6) = IIf(XMLparent.ChildNodes.Length <> 1, "", XMLchild.Text)
                ws.Cells(writeLine, 7) = XMLchild.nodeTypeString

                '�����o��
                For Each memberAttribute In XMLchild.Attributes
                    attributesWriteCol = attributesWriteCol + 1
                    ws.Cells(writeLine, attributesWriteCol) = memberAttribute.Name & "�F" & memberAttribute.Value
                Next
                '�����̐��ɂ���ăC���N�������g������������������
                attributesWriteCol = 7
                '==xml�t�@�C���̍������Č����ďo���\�[�X===========
                '�������R�����g�C������ꍇ�A�ʏ�o�͕����̓R�����g�A�E�g����B�Ȃ������o�͂͑ΏۊO
'                ws.Cells(writeLine, level) = XMLchild.nodeName & IIf(XMLparent.ChildNodes.Length <> 1, "", "�F" & XMLchild.Text)
                '==================================================
                Call writeXMLElement(XMLchild)
                '�ċA���������i���ɖ߂邽�ߊK�w���x������߂�
                level = level - 1
            End If
        Next
    End If
End Sub
