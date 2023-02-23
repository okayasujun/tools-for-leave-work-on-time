Attribute VB_Name = "try"
'#�V�F�C�v�𐮗񂳂���iY���W�̏����j
'�����m�̖��F���������̉摜������ƃG���[���o���B
Sub E_�������ɐ��񂳂���() 'TODO:����Ɋւ��Ă̓V���v���Ƀ��t�@�N�^�����O�����ׂ��B���ʕ����Ƃ��A���ꂢ�ɂ��悤�B
Attribute E_�������ɐ��񂳂���.VB_ProcData.VB_Invoke_Func = "t\n14"

    '���摜�Ԃ̊Ԋu
    Const MARGIN_BOTTOM = 70

    '�摜�̐����擾�ihttps://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype�j
    Dim pictureCount As Integer: pictureCount = 0
    For Each moveShape In ActiveSheet.Shapes
        If moveShape.Type = msoPicture Or moveShape.Type = msoGroup Then
            pictureCount = pictureCount + 1
        End If
    Next

    '�摜����Y���㕔���W���}�b�v�ŕێ�
    '�c�[�����Q�Ɛݒ肩��uMicrosoft Scripting Runtime�v�̗L�������K�v
    Dim shapeDic As Dictionary
    Set shapeDic = CreateObject("Scripting.Dictionary")
    
    Dim count As Integer: count = 0
    Dim shapeYArray() As Double
    
    '�}�b�v�ɏ���ݒ�A���W�\�[�g�p�ɍ��W���̔z��i�[
    For Each moveShape In ActiveSheet.Shapes
        If moveShape.Type = msoPicture Or moveShape.Type = msoGroup Then
        
            If count = 0 Then
                ReDim shapeYArray(count)
                shapeYArray(count) = moveShape.top
            Else
                ReDim Preserve shapeYArray(count)
                shapeYArray(count) = moveShape.top
            End If
            
            shapeDic.Add moveShape.Name, moveShape.top
            count = count + 1
        End If
    Next
    
    'top�v���p�e�B�̏��Ƀ\�[�g
    shapeYArray = sort(shapeYArray)
    
    '�㑱�����ɂ����ē��������̉摜������ꍇ�A�L�[�d���G���[�ɂȂ�̂�catch����
    On Error GoTo ErrHndl
    
    '�\�[�g��̃V�F�C�v�����l�߂Ȃ���
    Dim sortedShapeDic As Dictionary
    Set sortedShapeDic = CreateObject("Scripting.Dictionary")

    For Each yPoint In shapeYArray
        For Each dKey In shapeDic
            If yPoint = shapeDic.Item(dKey) Then
                sortedShapeDic.Add dKey, shapeDic.Item(dKey)
                Exit For
            End If
        Next
    Next

    '============================== �ȍ~�̏�����lineUpShapesOrderOfPasted()�Ɠ��� =================================

    '�ړ��ʒu���擾���邽�߂̃_�~�[�V�F�C�v
    Dim dummyShape As shape
    '�\�t���W���i�[����itop�͓s�x���������Aleft�͏����l���g���܂킷�j
    Dim top As Integer: top = Selection.top + 5
    Dim left As Integer: left = Selection.left
    '�L���v�V�������L�ڂ���p�̃Z��
    Dim captionRange As Range

    For Each dKey In sortedShapeDic
    
        Set moveShape = ActiveSheet.Shapes(dKey)
        
        '������̃Z�����擾���邽�߂̃_�~�[�V�F�C�v
        Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, left, top, 1, 1)
        
        '�V�F�C�v���ړ�����
        moveShape.top = dummyShape.TopLeftCell.top
        moveShape.left = left
        
        '�L���v�V�������͗p�Z�����擾����
        Set captionRange = dummyShape.TopLeftCell
        
        '�p�ς݂�����폜����
        dummyShape.Delete
        
        '���L���v�V�������͗p
        'Call setCaption(captionRange)
        
        '���Ώۂɂ����V�F�C�v�̏㕔���W + ���Ώۂɂ����V�F�C�v�̍��� + �摜�Ԃ̊Ԋu = ���̃V�F�C�v�̈ړ���㕔���W
        top = top + moveShape.height + MARGIN_BOTTOM
    Next
    Exit Sub
    
ErrHndl:
    MsgBox "���������̉摜�����邩��A�������炵�ă��g���C���ĂˁB"
End Sub
'�\�[�g�J�n
Function sort(ByRef targetArray() As Double)
    Dim swap As Double
    '�\�[�g�J�n
    For i = LBound(targetArray) To UBound(targetArray)
        For j = UBound(targetArray) To i Step -1
            If targetArray(i) > targetArray(j) Then
                swap = targetArray(i)
                targetArray(i) = targetArray(j)
                targetArray(j) = swap
            End If
        Next j
    Next i
    sort = targetArray
End Function
'============================================================================================================================
'#TODO:�K�v���E�g������̊ϓ_�ŗv�������B�O���[�v���Ώ۔͈͂��V�F�C�v���Ƃ��邩�I���Z���͈͓��Ƃ��邩�B
'#�I�𒆂̑�g���ɂ���V�F�C�v���O���[�v������B�O���[�v�������A�͂݃V�F�C�v�͍폜����
Sub M_�I�𒆘g���̃V�F�C�v�Q���O���[�v������()
    '�O���[�v���V�F�C�v�����J���}��؂�ŕێ�����p
    Dim targetShapeName As String
    '�J���}��؂�ŕێ��������̂�z���Ԃŕێ�����悤
    Dim targetShapeArray As Variant
    
    For Each shape In ActiveSheet.Shapes
        '�����Q�l�Fhttps://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype

        If shape.Type = msoAutoShape Or shape.Type = msoGroup Or shape.Type = msoPicture Then
            '��ӁA���ӁA�E�ӁA���ӂ���g���ɂ���V�F�C�v�݂̂�ΏۂƂ���
            If Selection.left < shape.left _
                And Selection.top < shape.top _
                And shape.left + shape.width < Selection.left + Selection.width _
                And shape.top + shape.height < Selection.top + Selection.height Then
                '�O���[�v�ΏۃV�F�C�v�̋L�^�i�㑱�����ŃO���[�v���j
                targetShapeName = targetShapeName & shape.Name & ","
            End If
        End If
        
    Next
    '�ΏۃV�F�C�v���͂��Ă����V�F�C�v���폜����
    Selection.Delete
    '�O���[�v�ΏۃV�F�C�v����z��
    targetShapeArray = Split(targetShapeName, ",")
    
    For Each shape In ActiveSheet.Shapes
        '�S�V�F�C�v�̒�����O���[�v�Ώۂ̂��̂����I����Ԃɂ���
        If isExistArray(targetShapeArray, shape.Name) Then
            shape.Select Replace:=False
        End If
    Next
    
    On Error GoTo catch
    
    If VarType(Selection) = vbObject Then
        '�I�𒆃V�F�C�v���O���[�v��
        Selection.Group.Select
    End If
    
    Exit Sub
catch:
End Sub
'#�z����ɑ��݂��邩�ǂ���
Function isExistArray(targetArray As Variant, checkValue As String)
    isExistArray = False
    
    If UBound(targetArray) = -1 Then
        'UBound�̖߂�l�F-1�͗v�f��0�������B���̏ꍇ�A���ׂđΏۊO�Ƃ���
        isExistArray = False
        Exit Function
    End If
    
    For i = LBound(targetArray) To UBound(targetArray)
        If targetArray(i) = checkValue Then
            isExistArray = True
            Exit For
        End If
    Next
End Function
'============================================================================================================================
