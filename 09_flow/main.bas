Attribute VB_Name = "main"
Sub createFlow()
    Const CENTER = 450
    '��0.��ʒu�w��A�V�F�C�v�\�擾
    
    '��1.�V�F�C�v���쐬����
    '���O�A�o���i�`�̃p�[�g�͂ǂ����ʂɗp�ӂ��悤�j
    '��2.�ꏊ�ړ�
    '��3.�R�l�N�^�t�^
    '
    '
    '
    '
    '
    '
    '
    Dim onShape As Shape
    Dim hjShape As Shape
    Dim shapeList As Variant: shapeList = Sheets("�V�F�C�v�\�Q").Range("A2:F31")
    Dim shapeType As Integer
    Dim lastRow As Integer: lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Dim cellColorCode As Long
    'Dim baseCenterPoint As Integer: baseCenterPoint = CENTER
    '��1.�V�F�C�v�쐬
    For i = 8 To lastRow
        '�t���[�쐬
        shapeType = vlookup(shapeList, Cells(i, 3), 2, 4)
        cellColorCode = Cells(i, 2).Interior.Color
        Set onShape = ActiveSheet.Shapes.AddShape(shapeType, 400, 100, 100, 30)
        onShape.TextFrame.Characters.Text = Cells(i, 4)
        onShape.Name = Cells(i, 2)
        Set onShape = baseShape(onShape)
        onShape.Fill.ForeColor.RGB = RGB(cellColorCode Mod 256, Int(cellColorCode / 256) Mod 256, Int(cellColorCode / 256 / 256))
        
        '�J�X�^��������΂����ɕ��������
        If Cells(i, 3) = "���[�v�J�n" Then
            onShape.AutoShapeType = 156
            onShape.Adjustments.Item(1) = 0.3
            onShape.Adjustments.Item(2) = 0
        ElseIf Cells(i, 3) = "���[�v�I��" Then
            onShape.AutoShapeType = 156
            onShape.Adjustments.Item(1) = 0
            onShape.Adjustments.Item(2) = 0.3
        ElseIf Cells(i, 3) = "�Q��" Then
            onShape.height = 30
            onShape.width = 30
        End If
    Next
    
    '��2.�ꏊ�ړ�
    Dim moveShape As Shape
    Dim baseTopPoint As Integer: baseTopPoint = 100
    Dim top As Integer: top = 100
    Dim nextShape As String
    Dim nextShapeArray As Variant
    Dim nextNo As String
    Dim movedShape As String
    For i = 7 To lastRow
        If i = 7 Then
            '�����_�T���ȁB�čl���悤�B
            nextShape = CStr(1)
        Else
            nextShape = CStr(Cells(i, 11))
        End If
        
        If nextShape Like "*" & vbLf & "*" Then
            '��������ꍇ
            nextShapeArray = Split(nextShape, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
                '�J�ڐ�̃O���[�v�i���������т�move��������Ⴄ���j
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                Set moveShape = ActiveSheet.Shapes(nextNo)
                '��������
                moveShape.Left = CENTER - moveShape.width / 2
                moveShape.top = top
                movedShape = movedShape & nextNo & ","
                
                moveShape.Left = CENTER - (UBound(nextShapeArray) - j) * moveShape.width * 1.3
                '�������ŏ����ȃV�F�C�v��������Ă݂邩
                Set hjShape = ActiveSheet.Shapes.AddShape(shapeType, 40, 10, 10, 10)
                hjShape.TextFrame.Characters.Text = Split(CStr(nextShapeArray(j)), ":")(0)
                Set hjShape = hojoShape(hjShape)
                hjShape.Left = moveShape.Left - 20
                hjShape.top = moveShape.top - 20
            Next
            top = top + moveShape.height + 25
        ElseIf Not movedShape Like "*" & nextShape & "*" Then
            '�P��ł��A�������������Ƃ������s��
            Set moveShape = ActiveSheet.Shapes(nextShape)
            '��������
            moveShape.Left = CENTER - moveShape.width / 2
            moveShape.top = top
            movedShape = movedShape & nextShape & ","
            top = top + moveShape.height + 25
        End If
    Next
    
    '��3.�R�l�N�^�t�^
    Dim beforeShape As Shape
    Dim afterShape As Shape
    Dim beforeShapeName As String
    Dim afterShapeName As String
    Dim connectShape As Shape
    Dim startPoint As Integer
    Dim endPoint As Integer
    Dim elbowNo As String
    For i = 8 To lastRow
        'before�V�F�C�v
        beforeShapeName = CStr(Cells(i, 2))
        Set beforeShape = ActiveSheet.Shapes(beforeShapeName)
        startPoint = vlookup(shapeList, beforeShape.AutoShapeType, 4, 5)
        
        
        'after�V�F�C�v
        afterShapeName = CStr(Cells(i, 11))
        If afterShapeName Like "*" & vbLf & "*" Then
            nextShapeArray = Split(afterShapeName, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
                '��������������
                Set connectShape = setConnect(connectShape)
                
                connectShape.ConnectorFormat.BeginConnect beforeShape, startPoint
                '�J�ڐ�̃O���[�v�i���������т�move��������Ⴄ���j
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                '������before�V�F�C�v������n�Ȃ�G���{�[�R�l�N�^
                If beforeShape.AutoShapeType = 63 Then
                    connectShape.ConnectorFormat.Type = msoConnectorElbow
                    elbowNo = elbowNo & nextNo & ","
                End If
                Set afterShape = ActiveSheet.Shapes(nextNo)
                endPoint = vlookup(shapeList, afterShape.AutoShapeType, 4, 6)
                connectShape.ConnectorFormat.EndConnect afterShape, endPoint
            Next
        ElseIf i <> lastRow Then
            Set afterShape = ActiveSheet.Shapes(afterShapeName)
            
            Set connectShape = setConnect(connectShape)
            connectShape.ConnectorFormat.BeginConnect beforeShape, startPoint
            endPoint = vlookup(shapeList, afterShape.AutoShapeType, 4, 6)
            connectShape.ConnectorFormat.EndConnect afterShape, endPoint
            '�������ɂ���
            If elbowNo Like "*" & beforeShapeName & "," & "*" Then
                connectShape.ConnectorFormat.Type = msoConnectorElbow
            End If
        End If
    
    Next
    
End Sub
Function setConnect(connectShape As Shape)
    '�R�l�N�g�V�F�C�v
    Set connectShape = ActiveSheet.Shapes.AddConnector( _
        Type:=1, _
        BeginX:=10, _
        BeginY:=1000, _
        EndX:=10, _
        EndY:=90)
    '���I�_�R�l�N�^���O�p�ɁB
    connectShape.Line.EndArrowheadStyle = msoArrowheadTriangle
    '�����̐F
    connectShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    '�����̑���
    connectShape.Line.Weight = 1
    Set setConnect = connectShape
End Function
'�\����g�p�������V�F�C�v��T��
Function vlookup(list As Variant, searchVal As String, searchCol As Integer, returnCol As Integer)
    vlookup = "61"
    For i = LBound(list) To UBound(list)
        If list(i, searchCol) Like "*" & searchVal & "*" Then
            vlookup = list(i, returnCol)
            Exit For
        End If
    Next
    
End Function
'�R�l�N�g�̎�ނ𒼐��ƃG���{�[�Ő؂�ւ���
Sub connectChange()
    If Selection.ShapeRange.ConnectorFormat.Type = msoConnectorElbow Then
        Selection.ShapeRange.ConnectorFormat.Type = msoConnectorStraight
    Else
        Selection.ShapeRange.ConnectorFormat.Type = msoConnectorElbow
    End If
End Sub
'�����I�������V�F�C�v�̂����A�ŏ��̃V�F�C�v�̍����W�ɑ��̃V�F�C�v�̈ʒu�����킹��
Sub leftPointSameSet()
    Dim baseLeftPoint As Integer: baseLeftPoint = Selection.ShapeRange.Item(1).Left
    For Each sp In Selection.ShapeRange
        sp.Left = baseLeftPoint
    Next
End Sub
'�I�������R�l�N�^�̎n�_���̃V�F�C�v�Ƃ̐ڑ��ʒu��ύX����
Sub checkExistsConnector()
    Dim currentBeginConnectPoint As Integer
    Dim targetShape As Shape
    
    '���݂̐ڑ��|�C���g�擾
    currentBeginConnectPoint = Selection.ShapeRange.ConnectorFormat.BeginConnectionSite
    '�n�_���̃V�F�C�v�擾
    Set targetShape = Selection.ShapeRange.ConnectorFormat.BeginConnectedShape
    '���̐ڑ��|�C���g�ɕύX
    If currentBeginConnectPoint = targetShape.ConnectionSiteCount Then
        currentBeginConnectPoint = 1
    Else
        currentBeginConnectPoint = currentBeginConnectPoint + 1
    End If
    '�ڑ��|�C���g��ύX
    Selection.ShapeRange.ConnectorFormat.BeginConnect targetShape, currentBeginConnectPoint
End Sub
Function baseShape(onShape As Shape)
    '���h��Ԃ��imsoTrue:����AmsoFalse:�Ȃ��j
    onShape.Fill.Visible = msoTrue
    '�����̑����B���D�݂łǂ���
    onShape.Line.Weight = 1
    '���F�w��
    onShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    '���h��Ԃ��F�w��
    onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    '�������F
    onShape.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    '������
    onShape.TextFrame.Characters.Font.Name = "Meiryo UI"
    '����������
    onShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    '����������
    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        '�e�L�X�g�ݒ��ɂ��Ȃ��Ƃ����Ȃ�
        onShape.TextFrame2.WordWrap = msoFalse
        onShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        If onShape.width < 100 Then
            onShape.TextFrame2.AutoSize = msoAutoSizeNone
            onShape.width = 100
        End If
        onShape.height = 30
    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    Set baseShape = onShape
End Function
Function hojoShape(onShape As Shape)
    '���h��Ԃ��imsoTrue:����AmsoFalse:�Ȃ��j
    onShape.Fill.Visible = msoFalse
    '�����̑����B���D�݂łǂ���
    onShape.Line.Weight = 0
    onShape.Line.Visible = msoFalse
    '���F�w��
    onShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    '���h��Ԃ��F�w��
    'onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    '�������F
    onShape.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    '������
    onShape.TextFrame.Characters.Font.Name = "Meiryo UI"
    '����������
    onShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    '����������
    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        '�e�L�X�g�ݒ��ɂ��Ȃ��Ƃ����Ȃ�
        onShape.TextFrame2.WordWrap = msoFalse
        onShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        If onShape.width < 30 Then
            onShape.TextFrame2.AutoSize = msoAutoSizeNone
            onShape.width = 30
        End If
        onShape.height = 15
    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    Set hojoShape = onShape
End Function
