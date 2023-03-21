Attribute VB_Name = "main"
Const PROCESS_NO_COL = 1
Const FLOW_NO_COL = 2
Const AFTER_NO_COL = 3
Const TYPE_COL = 4
Const FLOW_TEXT_COL = 5

Const LOOP_START_ROW = 2
Const HEIGHT_MARGIN = 32
Const WIDTH_MARGIN = 30
Dim LAST_FLOW_NO As String
Dim center As Integer
Dim shapeList As Variant
Dim flowList As Variant
'�쐬�V�F�C�v
Dim onShape As Shape
   
'�W���V�F�C�v
Dim hjShape As Shape
Function adjust()
    Dim currentShape As String
    Dim nextShape As String
    Dim nextShapeArray As Variant
    Dim beforeShape As Shape
    Dim moveShape As Shape
    Dim switchTotalWidth As Integer
    Dim TOP_CENTER As Integer: TOP_CENTER = ActiveSheet.Shapes("1").Left + ActiveSheet.Shapes("1").Width / 2
    '�V�F�C�v�̃^�C�v��\������
    Dim shapeType As String
    For i = LBound(flowList) To UBound(flowList)
        shapeType = CStr(flowList(i, TYPE_COL))
        '�J�ڐ���̎擾
        currentShape = CStr(flowList(i, FLOW_NO_COL))
        Set beforeShape = ActiveSheet.Shapes(currentShape)
        nextShape = CStr(flowList(i, AFTER_NO_COL))
        
        If nextShape Like "*" & vbLf & "*" Then
            '�J�ڐ悪��������ꍇ
            nextShapeArray = Split(nextShape, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
            
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                
                Set moveShape = ActiveSheet.Shapes(nextNo)
                
                If UBound(nextShapeArray) = 1 And Not shapeType Like "���[�v*" Then
                    If j = 0 Then
                        moveShape.top = beforeShape.top
                        moveShape.Left = beforeShape.Left - moveShape.Width - WIDTH_MARGIN
                        ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.BeginConnect beforeShape, 2
                        ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.EndConnect moveShape, 4
                        ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.Type = msoConnectorStraight
                        
'                        ActiveSheet.Shapes(moveShape.Name & "�̕⏕").top = ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).top + 2
'                        ActiveSheet.Shapes(moveShape.Name & "�̕⏕").Left = ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).Left + 2
                    ElseIf j = 1 Then
                        '�O�V�F�C�v�ƃZ���^�[�����킹��
                        Call setCenterPosition(moveShape, beforeShape.Left + beforeShape.Width / 2)
                        ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.Type = msoConnectorStraight
'                        ActiveSheet.Shapes(moveShape.Name & "�̕⏕").top = ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).top + 2
'                        ActiveSheet.Shapes(moveShape.Name & "�̕⏕").Left = ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).Left + 2
                    End If
                
                ElseIf UBound(nextShapeArray) > 1 Then
                    For k = LBound(nextShapeArray) To UBound(nextShapeArray)
                        nextNo = Split(CStr(nextShapeArray(k)), ":")(1)
                        switchTotalWidth = switchTotalWidth + ActiveSheet.Shapes(nextNo).Width + WIDTH_MARGIN
                    Next
                    a = switchTotalWidth / (UBound(nextShapeArray)) * (j)
                    'moveShape.Left = a
                    moveShape.Left = TOP_CENTER - switchTotalWidth / 2 + a - (moveShape.Width / 2)
                    switchTotalWidth = 0
                        ActiveSheet.Shapes(moveShape.Name & "�̕⏕").top = ActiveSheet.Shapes(moveShape.Name).top - 16
                        ActiveSheet.Shapes(moveShape.Name & "�̕⏕").Left = ActiveSheet.Shapes(moveShape.Name).Left + ActiveSheet.Shapes(moveShape.Name).Width / 2
                End If
            
                '�J�ڐ�̃O���[�v
            Next
        Else
            '�J�ڐ悪�P��
            Set moveShape = ActiveSheet.Shapes(nextShape)
            If nextShape = LAST_FLOW_NO Then
                '�ڑ��n�_��ς���
                ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.BeginConnect beforeShape, 2
                    ActiveSheet.Shapes(beforeShape.Name & "-" & moveShape.Name).ConnectorFormat.EndConnect moveShape, 2
            End If
        End If
    Next
    '���̕ӂŏI���v�f�̃R�l�N�^���œK���ł��邩�H
End Function
Function init()

    '�Z���^�[�|�W�V�����B�萔�����Ȃ̂ő啶����`�ɂ���
     center = Selection.Left + Selection.Width / 2
    
    '�g�p����V�F�C�v�̃��X�g
    shapeList = Sheets("�V�F�C�v�ꗗ").Range("A2:F31")
    
    '�V�F�C�v�̃^�C�v��\������
    Dim shapeType As Integer
    
    '�ŏI�s�i�t���[No�Ŏ擾����j
    Dim lastRow As Integer: lastRow = Sheets("�ڍ�").Cells(Rows.Count, 2).End(xlUp).Row
    
    '�t���[�̍ŏI�ԍ�
    LAST_FLOW_NO = Sheets("�ڍ�").Cells(lastRow, 2)
    
    '�g�p����V�F�C�v�̃��X�g
    flowList = Sheets("�ڍ�").Range("A2:R" & lastRow)
End Function
Sub A_�t���[�쐬()
Attribute A_�t���[�쐬.VB_ProcData.VB_Invoke_Func = "q\n14"

    Call init

    '��1.�V�F�C�v�쐬
    Call createFlowParts
    
    '��2.�ꏊ�ړ�
    Call moveFlowParts

    '��3.�R�l�N�^�t�^
    Call addConnector
    
    '��4.����
    Call adjust

End Sub
Function createFlowParts()
    '�����ΏۃZ���̐F�R�[�h
    Dim cellColorCode As Long
    '��1.�V�F�C�v�쐬
    For i = LBound(flowList) To UBound(flowList)
        '�V�F�C�v��ʎ擾
        shapeType = vlookup(shapeList, flowList(i, TYPE_COL), 2, 4)
        '�t���[No��̃Z���F���擾
        cellColorCode = 16777215 'flowList(i, 2).Interior.Color
        '�V�F�C�v�𐶐�
        Set onShape = ActiveSheet.Shapes.AddShape(shapeType, 400, 100, 100, 30)
        '�e�L�X�g�̐ݒ�
        onShape.TextFrame.Characters.Text = flowList(i, FLOW_TEXT_COL)
        '���O�̐ݒ�
        onShape.Name = flowList(i, 2)
        '�V�F�C�v����
        Set onShape = baseShape(onShape)
        '�V�F�C�v�̐F�ݒ�
        onShape.Fill.ForeColor.RGB = RGB(cellColorCode Mod 256, Int(cellColorCode / 256) Mod 256, Int(cellColorCode / 256 / 256))
        
        '�`�̒���������΂����ɕ��������
        If flowList(i, TYPE_COL) = "���[�v�J�n" Then
            onShape.AutoShapeType = 156
            onShape.Adjustments.Item(1) = 0.3
            onShape.Adjustments.Item(2) = 0
        ElseIf flowList(i, TYPE_COL) = "���[�v�I��" Then
            onShape.AutoShapeType = 156
            onShape.Adjustments.Item(1) = 0
            onShape.Adjustments.Item(2) = 0.3
        ElseIf flowList(i, TYPE_COL) = "�Q��" Then
            onShape.Height = 30
            onShape.Width = 30
        End If
    Next
End Function
Function moveFlowParts()
    '��2.�ꏊ�ړ�
    Dim moveShape As Shape
    Dim top As Integer: top = Selection.top '100
    Dim nextShape As String
    Dim nextShapeArray As Variant
    '
    Dim nextNo As String
    '�ړ��ς݃V�F�C�v�̖��̂��J���}��؂�ŊǗ�
    Dim movedShape As String
    
    '�J�n�v�f�̈ړ�
    Set moveShape = ActiveSheet.Shapes("1")
    '��������
    moveShape.Left = center - moveShape.Width / 2
    moveShape.top = top
    movedShape = movedShape & nextShape & ","
    top = top + moveShape.Height + HEIGHT_MARGIN
    
    For i = LBound(flowList) To UBound(flowList)
        '�J�ڐ���̎擾
        nextShape = CStr(flowList(i, AFTER_NO_COL))
        
        '���[�v�I���v�f��������END�J�ڂŒP�ꏈ���̃��W�b�N�������ɒʂ点��
        If flowList(i, TYPE_COL) = "���[�v�I��" Then
            nextShape = Split(Split(nextShape, vbLf)(1), ":")(1)
        End If
        
        If nextShape Like "*" & vbLf & "*" Then
            '�J�ڐ悪��������ꍇ
            nextShapeArray = Split(nextShape, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
                '�J�ڐ�̃O���[�v
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                
                Set moveShape = ActiveSheet.Shapes(nextNo)
                    
                '��������
                moveShape.Left = center - moveShape.Width / 2
                moveShape.top = top
                movedShape = movedShape & nextNo & ","
                '�����������тɂ����邽��
                moveShape.Left = center - (UBound(nextShapeArray) - j) * moveShape.Width * 1.3
                
                '�ē��e�L�X�g�̃V�F�C�v
                Set hjShape = ActiveSheet.Shapes.AddShape(61, 40, 10, 10, 10)
                hjShape.TextFrame.Characters.Text = Split(CStr(nextShapeArray(j)), ":")(0)
                hjShape.Name = moveShape.Name & "�̕⏕"
                Set hjShape = hojoShape(hjShape)
                hjShape.Left = moveShape.Left - 30
                hjShape.top = moveShape.top - 20
            Next
            '����top�v���p�e�B�ݒ�
            top = top + moveShape.Height + HEIGHT_MARGIN
        ElseIf Not movedShape Like "*" & nextShape & "*" Then
            '�P��ł��A�������������Ƃ������s��
            Set moveShape = ActiveSheet.Shapes(nextShape)
            '��������
            moveShape.Left = center - moveShape.Width / 2
            moveShape.top = top
            movedShape = movedShape & nextShape & ","
            top = top + moveShape.Height + HEIGHT_MARGIN
        End If
    Next
    '�I���v�f�̈ړ�
    Set moveShape = ActiveSheet.Shapes(LAST_FLOW_NO)
    '��������
    moveShape.Left = center - moveShape.Width / 2
    moveShape.top = top
End Function
Function addConnector()
    '��3.�R�l�N�^�t�^
    Dim beforeShape As Shape
    Dim afterShape As Shape
    Dim beforeShapeName As String
    Dim afterShapeName As String
    Dim connectShape As Shape
    Dim startPoint As Integer
    Dim endPoint As Integer
    '�G���{�[�R�l�N�^���n�_�ɂ��Ă�V�F�C�v���̂��J���}��؂�ŊǗ�
    Dim elbowNo As String
    For i = LBound(flowList) To UBound(flowList) - 1
        'before�V�F�C�v
        beforeShapeName = CStr(flowList(i, FLOW_NO_COL))
        Set beforeShape = ActiveSheet.Shapes(beforeShapeName)
        startPoint = vlookup(shapeList, beforeShape.AutoShapeType, 4, 5)
        
        'after�V�F�C�v
        afterShapeName = CStr(flowList(i, AFTER_NO_COL))
        If afterShapeName Like "*" & vbLf & "*" Then
            '�J�ڐ悪��������ꍇ
            nextShapeArray = Split(afterShapeName, vbLf)
            For j = LBound(nextShapeArray) To UBound(nextShapeArray)
                '�R�l�N�^����
                Set connectShape = setConnect(connectShape)
                
                '�R�l�N�^�ڑ��i�n�_�j
                connectShape.ConnectorFormat.BeginConnect beforeShape, startPoint
                
                '�J�ڐ�̃O���[�v�i���������т�move��������Ⴄ���j
                nextNo = Split(CStr(nextShapeArray(j)), ":")(1)
                
                '������before�V�F�C�v������n�Ȃ�G���{�[�R�l�N�^
                If beforeShape.AutoShapeType = 63 Or beforeShape.AutoShapeType = 156 Then
                    connectShape.ConnectorFormat.Type = msoConnectorElbow
                    elbowNo = elbowNo & nextNo & ","
                End If
                
                '�ڑ���V�F�C�v�̎擾
                Set afterShape = ActiveSheet.Shapes(nextNo)
                '
                endPoint = vlookup(shapeList, afterShape.AutoShapeType, 4, 6)
                '�R�l�N�^�ڑ��i�I�_�j
                connectShape.ConnectorFormat.EndConnect afterShape, endPoint
                
                '�u���[�v�I���v����u�X�[�v�J�n�v�ɐL�т�R�l�N�^�̏ꍇ
                If Split(CStr(nextShapeArray(j)), ":")(0) = "Next" Then
                    connectShape.ConnectorFormat.BeginConnect beforeShape, 1
                    connectShape.ConnectorFormat.EndConnect afterShape, 1
                End If
                connectShape.Name = beforeShape.Name & "-" & afterShape.Name
            Next
        ElseIf i <> lastRow Then
            '�ڑ��悪�P��ōŏI�s�ł��Ȃ��ꍇ
            Set afterShape = ActiveSheet.Shapes(afterShapeName)
            
            Set connectShape = setConnect(connectShape)
            connectShape.ConnectorFormat.BeginConnect beforeShape, startPoint
            endPoint = vlookup(shapeList, afterShape.AutoShapeType, 4, 6)
            connectShape.ConnectorFormat.EndConnect afterShape, endPoint
            '�������ɂ���
            If elbowNo Like "*" & beforeShapeName & "," & "*" Then
                connectShape.ConnectorFormat.Type = msoConnectorElbow
            End If
            connectShape.Name = beforeShape.Name & "-" & afterShape.Name
        End If
    
    Next
End Function
Function setCenterPosition(shp As Shape, center)
    shp.Left = center - shp.Width / 2
End Function
Function setConnect(connectShape As Shape)
    '�R�l�N�g�V�F�C�v
    Set connectShape = ActiveSheet.Shapes.addConnector( _
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
Function vlookup(list As Variant, searchVal As Variant, searchCol As Integer, returnCol As Integer)
    vlookup = "61"
    For i = LBound(list) To UBound(list)
        If list(i, searchCol) Like "*" & searchVal & "*" Then
            vlookup = list(i, returnCol)
            Exit For
        End If
    Next
    
End Function

'�R�l�N�^�̎�ނ𒼐��ƃG���{�[�Ő؂�ւ���
Sub B_�R�l�N�^�ؑ֒����G���{�[()
Attribute B_�R�l�N�^�ؑ֒����G���{�[.VB_ProcData.VB_Invoke_Func = "w\n14"

    If connectorErrorCheck Then
       Exit Sub
    End If
    
    With Selection.ShapeRange.ConnectorFormat
        If .Type = msoConnectorElbow Then
            .Type = msoConnectorStraight
        Else
            .Type = msoConnectorElbow
        End If
    End With
End Sub
'�����I�������V�F�C�v�̂����A�ŏ��̃V�F�C�v�̍����W�ɑ��̃V�F�C�v�̈ʒu�����킹��
'TODO:����{����left���W���킹����Ȃ���1�ڂ̃Z���^�[���킹����Ȃ��Ƃ������
'������E�E�EbaseXCenterPoint��baseYCenterPoint��2���~�����B�܂��ȒP�ɍ��邩
Sub C_X���W���킹()
Attribute C_X���W���킹.VB_ProcData.VB_Invoke_Func = "e\n14"
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).Left + Selection.ShapeRange.Item(1).Width / 2
    For Each sp In Selection.ShapeRange
        sp.Left = baseXCenterPoint - sp.Width / 2
    Next
End Sub
Sub D_Y���W���킹()
Attribute D_Y���W���킹.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).top + Selection.ShapeRange.Item(1).Height / 2
    For Each sp In Selection.ShapeRange
        sp.top = baseXCenterPoint - sp.Height / 2
    Next
End Sub
'�I�������R�l�N�^�̎n�_���̃V�F�C�v�Ƃ̐ڑ��ʒu��ύX����
Sub E_�R�l�N�^�n�_�ύX()
Attribute E_�R�l�N�^�n�_�ύX.VB_ProcData.VB_Invoke_Func = "r\n14"

    If connectorErrorCheck Then
       Exit Sub
    End If
    
    Dim currentBeginConnectPoint As Integer
    Dim targetShape As Shape
    
    With Selection.ShapeRange.ConnectorFormat
        '���݂̎n�_�ڑ��|�C���g�擾
        currentBeginConnectPoint = .BeginConnectionSite
        '�n�_���̃V�F�C�v�擾
        Set targetShape = .BeginConnectedShape
        '���̐ڑ��|�C���g�ɕύX�i���݂�MAX�̏ꍇ1�ɖ߂��j
        If currentBeginConnectPoint = targetShape.ConnectionSiteCount Then
            currentBeginConnectPoint = 1
        Else
            currentBeginConnectPoint = currentBeginConnectPoint + 1
        End If
        '�ڑ��|�C���g��ύX
        .BeginConnect targetShape, currentBeginConnectPoint
    End With
End Sub
'�I�������R�l�N�^�̏I�_���̃V�F�C�v�Ƃ̐ڑ��ʒu��ύX����
Sub F_�R�l�N�^�I�_�ύX()
Attribute F_�R�l�N�^�I�_�ύX.VB_ProcData.VB_Invoke_Func = "t\n14"
    If connectorErrorCheck Then
       Exit Sub
    End If
    
    Dim currentEndConnectPoint As Integer
    Dim targetShape As Shape
    
    With Selection.ShapeRange.ConnectorFormat
        '���݂̎n�_�ڑ��|�C���g�擾
        currentEndConnectPoint = .EndConnectionSite
        '�n�_���̃V�F�C�v�擾
        Set targetShape = .EndConnectedShape
        '���̐ڑ��|�C���g�ɕύX�i���݂�MAX�̏ꍇ1�ɖ߂��j
        If currentEndConnectPoint = targetShape.ConnectionSiteCount Then
            currentEndConnectPoint = 1
        Else
            currentEndConnectPoint = currentEndConnectPoint + 1
        End If
        '�ڑ��|�C���g��ύX
        .EndConnect targetShape, currentEndConnectPoint
    End With
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
    If onShape.Width < 100 Then
        onShape.TextFrame2.AutoSize = msoAutoSizeNone
        onShape.Width = 100
    End If
    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    onShape.TextFrame2.WordWrap = msoTrue
    onShape.TextFrame2.AutoSize = msoAutoSizeNone
    onShape.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
    onShape.Height = 30
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
    If onShape.Width < 30 Then
        onShape.TextFrame2.AutoSize = msoAutoSizeNone
        onShape.Width = 30
    End If
    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    onShape.Height = 15
    Set hojoShape = onShape
End Function
'�R�l�N�^�ɏ���������ۂ̃G���[�`�F�b�N�B�߂�lTrue�ŃG���[����AFalse�ŃG���[�Ȃ�
Function connectorErrorCheck()
    If TypeName(Selection) = "Range" Then
        MsgBox "�R�l�N�^��I��ł��������B"
        connectorErrorCheck = True
        Exit Function
    End If
    If Not Selection.ShapeRange.Connector Then
        MsgBox "�R�l�N�^��I��ł��������B"
        connectorErrorCheck = True
    End If
End Function
