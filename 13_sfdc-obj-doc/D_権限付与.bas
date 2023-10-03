Attribute VB_Name = "D_�����t�^"
Dim itemSheet As Worksheet
Sub D_�����t�^()

    Const PRE = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<Profile xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf
    Const SUF = "</Profile>"

'�����̔Ԃ͕ҏW�s��
'�K�{���ڂ͎Q�ƁE�ҏW���Z�ł��邱��
'�ȂǁA�ŏ��Ƀ`�F�b�N�������K�v
    
    Set itemSheet = Sheets(ITEM_SHEET)
    
    Dim objApiName As String: objApiName = itemSheet.Cells(2, 4)
    Dim filePath As String: filePath = ThisWorkbook.Path & "\profiles\Admin.profile-meta.xml"
    Dim fileName As String: fileName = filePath '�{���͂����ŁA�t�@�C�����ݒ肷��
    
    Dim lastRow As Integer: lastRow = itemSheet.Cells(4, 1).End(xlDown).row
    
    Dim itemApiName As String, readPermission As String, editPermission As String
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        
        .writeText PRE
        For i = 5 To lastRow
            If itemSheet.Cells(i, 17) = "�Z" Then
                GoTo continue
            End If
            itemApiName = itemSheet.Cells(i, 5).Value
            editPermission = itemSheet.Cells(i, 40) = "�Z"
            readPermission = itemSheet.Cells(i, 39) = "�Z"
            '38,39
            .writeText "    <fieldPermissions>" & vbCrLf
            .writeText "        <editable>" & editPermission & "</editable>" & vbCrLf
            .writeText "        <field>" & objApiName & "." & itemApiName & "</field>" & vbCrLf
            .writeText "        <readable>" & readPermission & "</readable>" & vbCrLf
            .writeText "    </fieldPermissions>" & vbCrLf
continue:
        Next
        '�^�u�̐ݒ�
        .writeText "    <tabVisibilities>" & vbCrLf
        .writeText "        <tab>" & objApiName & "</tab>" & vbCrLf
        .writeText "        <visibility>" & itemSheet.Cells(3, 40) & "</visibility>" & vbCrLf
        .writeText "    </tabVisibilities>" & vbCrLf
        
        .writeText SUF
        '�����o�������I��
        .Position = 0
        .Type = 1
        .Position = 3
        bytetmp = .Read
        .SaveToFile fileName, 2
        '�R�s�[��t�@�C�������
        .Close
    End With
    
    'UTF-8�Ńe�L�X�g�t�@�C���֏o�͂���
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .LineSeparator = 10
        .Type = 1
        .Open
        .write bytetmp
        .SetEOS
        .SaveToFile fileName, 2
        .Close
    End With
    
    MsgBox "�������܂����B"
End Sub
