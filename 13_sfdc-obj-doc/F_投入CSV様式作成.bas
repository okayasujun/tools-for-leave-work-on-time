Attribute VB_Name = "F_����CSV�l���쐬"
Dim itemSheet As Worksheet
Dim csvSheet As Worksheet
Sub F_����CSV�l���쐬()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    '�f�[�^���[�_�p��CSV�t�@�C����f���o���}�N���B�܂��������Ă�낤�E�E(���L�́M)��
    Set itemSheet = Sheets(ITEM_SHEET)
    
    Worksheets.Add
    '�V�[�g����ύX�i�ǉ����ꂽ�V�[�g�̓A�N�e�B�u�ƂȂ�j
    ActiveSheet.Name = "dataloader_format"
    Set csvSheet = ActiveSheet
    
    Dim writeCol As Integer: writeCol = 1
    With itemSheet
        For i = 5 To .Cells(4, 1).End(xlDown).row
            If .Cells(i, 2) = "�Z" And .Cells(i, 7) <> "�����̔�" And .Cells(i, 8) = "" Then
                  
                '���x����
                csvSheet.Cells(1, writeCol) = .Cells(i, 3).Value
                'API��
                csvSheet.Cells(2, writeCol) = .Cells(i, 5).Value
                '�f�[�^�^
                csvSheet.Cells(3, writeCol) = .Cells(i, 7).Value
                '�I�����X�g
                csvSheet.Cells(5, writeCol) = .Cells(i, 14).Value
                '�񕝒������܂������Ȃ�
                csvSheet.Cells.EntireColumn.AutoFit
                '�K�{�}�[�N
                If .Cells(i, 17) = "�Z" Then
                    csvSheet.Cells(4, writeCol) = "�K�{�I"
                End If
                If .Cells(i, 18) = "�Z" Then
                    csvSheet.Cells(4, writeCol) = csvSheet.Cells(4, writeCol) & "��ӁI"
                End If
                writeCol = writeCol + 1
            End If
        Next
    End With
    '���܂������Ȃ��񕝒���
    csvSheet.Cells.EntireColumn.AutoFit
    
    '�ۑ�
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Dim saveDir As String: saveDir = ThisWorkbook.Path & "\" & csvSheet.Name & Format(Now, "yyyyddmm-hhmmss") & ".csv"
    Sheets(csvSheet.Name).Copy
    ActiveWorkbook.SaveAs saveDir, FileFormat:=xlCSV, Local:=True
    ActiveWorkbook.Close
    csvSheet.Delete
    
    MsgBox "�������܂����B"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
