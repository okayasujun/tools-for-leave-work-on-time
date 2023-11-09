Attribute VB_Name = "C_�J�X�^�����ڃ��^�f�[�^�쐬"
Const PICK_LIST_PRE_TAG = "    " & "<valueSet>" & vbCrLf & _
                          "    " & "<restricted>true</restricted>" & vbCrLf & _
                          "    " & "<valueSetDefinition>" & vbCrLf & _
                          "    " & "    <sorted>false</sorted>" & vbCrLf
Const PICK_LIST_SUF_TAG = "    " & "        </valueSetDefinition>" & vbCrLf & _
                          "    " & "</valueSet>" & vbCrLf
Const PRE = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
            "<CustomField xmlns=""http://soap.sforce.com/2006/04/metadata"">" & vbCrLf
Const SUF = "</CustomField>"
'
'�J�X�^�����ڂ��ƂɃ��^�f�[�^�t�@�C�����쐬����
'�f�[�^�^�ɂ���Ē�`���K�v�ȃ^�O��[]�V�[�g�ɒ�`���Ă���
'
Sub D_�J�X�^�����ڃ��^�f�[�^�쐬()
    Call initiarize

    '�ŏI�s
    Dim lastRow As Integer: lastRow = itemSheet.Cells(4, 1).End(xlDown).row
    '���x�����AAPI��
    Dim labelName As String
    Dim apiName As String
    
    '###################################################
    '�����O�ɃG���[�`�F�b�N���K�v
    '�EAPI���̐擪���啶�������邱��
    '
    '###################################################
    
    Dim i As Integer
    For i = 5 To lastRow
        '�L�����ڂ̂ݑΏۂƂ���
        If itemSheet.Cells(i, 2).Value = "�Z" Then
 
            apiName = itemSheet.Cells(i, 5).Value
            fileName = fieldsDirPath & apiName & ".field-meta.xml"

            Call openStream
            '�����o�������J�n
            stream.writeText PRE, 0
            stream.writeText getItemMetaData(i), 0
            stream.writeText SUF, 0
            Call saveTextWithUTF8(stream, fileName)
        End If
    Next
    
    MsgBox "�������܂����B"
End Sub
'���^�f�[�^�̃e�L�X�g����Ԃ�
Function getItemMetaData(row As Integer)
    Dim writeText As String
    Dim returnValue As String
    Dim dataTypeColumn As Integer
    Dim valueColumn As Integer
    Dim openTag As String
    Dim closeTag As String
    Dim dataType As String: dimdataType = itemSheet.Cells(row, 7).Value
    Dim valueType As String
    Dim listArray As Variant
    Dim listOneArray As Variant
    Dim listFlag As Boolean
    '���^�f�[�^�t�@�C���ɏ����o�����ǂ������u�Z�v�̗L���Ŏ擾����
    Dim writeTagFlag As Boolean: writeTagFlag = True
    
    dataType = itemSheet.Cells(row, 7).Value
    dataType = IIf(itemSheet.Cells(row, 8).Value = "�Z", "(����)" & dataType, dataType)
    
    '�������s�̃f�[�^�^���`�V�[�g����T���i�E�������[�v�j
    For i = 4 To 31
        If dataType = itemMetaSheet.Cells(2, i).Value Then
            dataTypeColumn = i
            Exit For
        End If
    Next
    '�c�������[�v
    For i = 3 To 37
        valueColumn = itemMetaSheet.Cells(i, 2).Value
        
        If itemMetaSheet.Cells(i, dataTypeColumn).Value And valueColumn > 0 Then
            valueType = itemMetaSheet.Cells(i, 3).Value
            writeText = itemSheet.Cells(row, valueColumn).Value
            
            If valueType = "�e�L�X�g" Then
            
'                If itemMetaSheet.Cells(i, 1).Value = "<defaultValue>" And writeText <> "" Then
'                    writeText = "&quot;" & writeText & "&quot;" '������̂Ƃ��͂���Ȃ��E�E�E
'                End If
'                TODO:�f�t�H���g�l�̑Ή��͕K�v�i�����̌������炵�āj

            ElseIf valueType = "���l" Then
                
            ElseIf valueType = "�^�U" Then
                writeText = IIf(writeText = "�Z", "True", "False")
                
            ElseIf valueType = "���X�g" Then
                Debug.Print writeText
                listArray = Split(writeText, vbLf)
                listFlag = True
                writeTagFlag = False
            End If
            
            If writeTagFlag Then
                openTag = itemMetaSheet.Cells(i, 1).Value
                closeTag = Replace(openTag, "<", "</")
                returnValue = returnValue & "    " & openTag & writeText & closeTag & vbCrLf
            End If
            writeTagFlag = True
        End If
    Next
    
    '�I�����X�g�̃^�O�ݒ�
    If listFlag Then
        returnValue = returnValue & PICK_LIST_PRE_TAG
        For Each Item In listArray
            listOneArray = Split(Item, ":")
            returnValue = returnValue & "    " & "<value>" & vbCrLf
            returnValue = returnValue & "    " & "<fullName>" & listOneArray(0) & "</fullName>" & vbCrLf
            returnValue = returnValue & "    " & "<default>false</default>" & vbCrLf
            returnValue = returnValue & "    " & "<label>" & listOneArray(1) & "</label>" & vbCrLf
            returnValue = returnValue & "    " & "</value>" & vbCrLf
        Next
        returnValue = returnValue & "    " & "        </valueSetDefinition>" & vbCrLf
        returnValue = returnValue & "    " & "</valueSet>" & vbCrLf
    End If
    getItemMetaData = returnValue
End Function

