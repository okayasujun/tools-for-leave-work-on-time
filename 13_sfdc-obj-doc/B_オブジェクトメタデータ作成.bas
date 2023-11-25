Attribute VB_Name = "B_�I�u�W�F�N�g���^�f�[�^�쐬"
Sub A_�I�u�W�F�N�g���^�f�[�^�쐬()
    Call initiarize
    fileName = ThisWorkbook.path & "\objects\" & objApiName & "\" & objApiName & ".object-meta.xml"
    
    '�A�z�z��쐬�iA,D���Ή��t����j
    Dim objectInformation As Object
    Set objectInformation = CreateObject("Scripting.Dictionary")
    For i = 1 To objSheet.Cells(Rows.Count, 1).End(xlUp).row
        objectInformation(objSheet.Cells(i, 1).Value) = objSheet.Cells(i, 4).Value
    Next
    
    '���K�\���i�u���̂��߁j
    Call setupRegexp("{.*(?=})")
    
    '�t�@�C���ɏ����o���e�L�X�g
    Dim writeText As String
    '�u��������B���{}�̒��̕������i�[����
    Dim replaceValue As String
    '�A�z�z��ւ̃A�N�Z�X�L�[
    Dim key As String
    
    '�e�L�X�g�t�@�C���o�͏���
    Call openStream
    
    '�����o�������J�n
    For i = 1 To objMetaSheet.Cells(Rows.Count, 1).End(xlUp).row
        writeText = objMetaSheet.Cells(i, 1).Value
            
        '�g���ʂ�����ꍇ�͒u���������K�v�i�������K�\���`�F�b�N�ɂ������j
        If writeText Like "*{*" Then
            'Name���ڂ��e�L�X�g�^�ɂ���Ƃ���displayFormat�^�O�͕s�v
            If Not (writeText Like "*{�\���`��}*" And objectInformation("�f�[�^�^") = "Text") Then
            
                '��{�\�����x���iVBA�͍m���ǂ݂ł��Ȃ�����O���ʂ͎c��j
                replaceValue = regexp.Execute(writeText)(0)
                '��F�\�����x��
                key = Replace(replaceValue, "{", "")
                '��F<label>�\�����x��</label> �����Ȃ�Ƃ������[��
                writeText = Replace(writeText, "{", "")
                writeText = Replace(writeText, "}", "")
                '��F<label>hoge</label>
                writeText = Replace(writeText, key, objectInformation(key))
            End If
        End If
            
        stream.writeText writeText & vbCrLf
    Next
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "�������܂����B"
End Sub

Sub B_�^�u���^�f�[�^�쐬()
    Call initiarize
    fileName = ThisWorkbook.path & "\tabs\" & objApiName & ".tab-meta.xml"

    '�e�L�X�g�t�@�C���o�͏���
    Call openStream
    '�t�@�C���ɏ����o���e�L�X�g
    Dim writeText As String: writeText = _
        "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<CustomTab xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf & _
        "    <customObject>true</customObject>" & vbCrLf & _
        "    <motif>Custom93: Shopping Cart</motif>" & vbCrLf & _
        "</CustomTab>"

    stream.writeText writeText & vbCrLf
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "�������܂����B"
End Sub

Sub C_���X�g�r���[���^�f�[�^�쐬()
    Call initiarize
    fileName = ThisWorkbook.path & "\objects\" & objApiName & "\listViews\All.listView-meta.xml"

    Call openStream
    '�t�@�C���ɏ����o���e�L�X�g
    Dim writeText As String: writeText = _
        "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<ListView xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf & _
        "    <fullName>All</fullName>" & vbCrLf & _
        "    <filterScope>Everything</filterScope>" & vbCrLf & _
        "    <label>���ׂđI��</label>" & vbCrLf & _
        "</ListView>"

    stream.writeText writeText & vbCrLf
    Call saveTextWithUTF8(stream, fileName)
    MsgBox "�������܂����B"
End Sub

