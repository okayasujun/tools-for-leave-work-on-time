Attribute VB_Name = "B_�I�u�W�F�N�g���^�f�[�^�쐬"
'
'object�̃��^�f�[�^�t�@�C�����쐬����
'[CustomObject]�V�[�g�ɒ�`�������e���x�[�X�ɗv����u������
'
'
Sub B_�I�u�W�F�N�g���^�f�[�^�쐬()
    Dim objSheet As Worksheet
    Set objSheet = Sheets(OBJECT_SHEET)
    Dim objMetaSheet As Worksheet
    Set objMetaSheet = Sheets(OBJECT_META_SHEET)
    
    '�I�u�W�F�N�g��API��
    Dim objApiName As String: objApiName = Sheets(OBJECT_SHEET).Cells(4, 4).Value
    '�����Ώۃt�@�C�����B��uCustomObject04__c.object-meta.xml�v
    Dim fileName As String: fileName = ThisWorkbook.Path & "\objects\" & objApiName & "\" & objApiName & ".object-meta.xml"

    '�A�z�z��쐬�iA,D���Ή��t����j
    Dim objectInformation As Object
    Set objectInformation = CreateObject("Scripting.Dictionary")
    For i = 1 To objSheet.Cells(Rows.Count, 1).End(xlUp).row
        With objSheet
            objectInformation(.Cells(i, 1).Value) = .Cells(i, 4).Value
        End With
    Next
    
    '���K�\���i�u���̂��߁j
    Dim regexpObj As Object
    Set regexpObj = CreateObject("VBScript.RegExp")
    With regexpObj
        '�u���������o�p�p�^�[���iVBA�ōm���ǂ݂͎g���Ȃ��j
        .Pattern = "{.*(?=})"
        '�p�啶������������ʂ��Ȃ�
        .IgnoreCase = True
        '������S�̂ɑ΂��ăp�^�[���}�b�`������
        .Global = True
    End With
    '�t�@�C���ɏ����o���e�L�X�g
    Dim writeText As String
    '�u��������B���{}�̒��̕������i�[����
    Dim replaceValue As String
    '�A�z�z��ւ̃A�N�Z�X�L�[
    Dim key As String
    
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Charset = "UTF-8"
    st.Open
    
    '�����o�������J�n
    For i = 1 To objMetaSheet.Cells(Rows.Count, 1).End(xlUp).row
        writeText = objMetaSheet.Cells(i, 1).Value
            
        '�g���ʂ�����ꍇ�͒u���������K�v�i�������K�\���`�F�b�N�ɂ������j
        If writeText Like "*{*" Then
            'Name���ڂ��e�L�X�g�^�ɂ���Ƃ���displayFormat�^�O�͕s�v
            If Not (writeText Like "*{�\���`��}*" And objectInformation("�f�[�^�^") = "Text") Then
            
                '��{�\�����x���iVBA�͍m���ǂ݂ł��Ȃ�����O���ʂ͎c��j
                replaceValue = regexpObj.Execute(writeText)(0)
                '��F�\�����x��
                key = Replace(replaceValue, "{", "")
                '��F<label>�\�����x��</label>
                writeText = Replace(writeText, "{", "")
                writeText = Replace(writeText, "}", "")
                '��F<label>hoge</label>
                writeText = Replace(writeText, key, objectInformation(key))
            End If
        End If
            
        st.writeText writeText & vbCrLf
    Next
    Call saveTextWithUTF8(st, fileName)

    
    '�^�u�t�@�C���쐬
    If objectInformation("�^�u���쐬") = True Then
    
        fileName = ThisWorkbook.Path & "\tabs\" & objApiName & ".tab-meta.xml"

        Set st = CreateObject("ADODB.Stream")
        st.Charset = "UTF-8"
        st.Open
        writeText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                    "<CustomTab xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf & _
                    "    <customObject>true</customObject>" & vbCrLf & _
                    "    <motif>Custom93: Shopping Cart</motif>" & vbCrLf & _
                    "</CustomTab>"

        st.writeText writeText & vbCrLf
        Call saveTextWithUTF8(st, fileName)
    End If
    
    '���ׂĕ\���̃��X�g�r���[�쐬
    fileName = ThisWorkbook.Path & "\objects\" & objApiName & "\listViews\All.listView-meta.xml"

    Set st = CreateObject("ADODB.Stream")
    st.Charset = "UTF-8"
    st.Open
    writeText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<ListView xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf & _
                "    <fullName>All</fullName>" & vbCrLf & _
                "    <filterScope>Everything</filterScope>" & vbCrLf & _
                "    <label>���ׂđI��</label>" & vbCrLf & _
                "</ListView>"

    st.writeText writeText & vbCrLf
    Call saveTextWithUTF8(st, fileName)
    
    MsgBox "�������܂����B"
End Sub
'UTF-8�ŕۑ�����Ƃ��̕ۑ��������X�g���[��object�ƃt�@�C�����ōs��
Function saveTextWithUTF8(stream As Object, fileFullName As String)
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
        '�ۑ�
        stream.SaveToFile fileFullName, 2
        '�R�s�[��t�@�C�������
        stream.Close
End Function

