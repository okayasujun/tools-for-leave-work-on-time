Attribute VB_Name = "main"
'����No��
Const PROCESS_NO_COL = 1
'�t���[No��
Const FLOW_NO_COL = 2
'�J�ڐ��
Const DIST_NO_COL = 3
'��ʗ�
Const TYPE_COL = 4
'�V�F�C�v�e�L�X�g��
Const SHAPE_TEXT_COL = 5
'[�ڍ�]�V�[�g�̑����J�n�s
Const LOOP_START_ROW = 2
'�V�F�C�v�c�Ԋu
Const HEIGHT_MARGIN = 32
'�V�F�C�v���Ԋu
Const WIDTH_MARGIN = 130
'�V�F�C�v�̕�
Const SHAPE_WIDTH = 130
'�V�F�C�v�̍���
Const SHAPE_HEIGHT = 33
'�A�j���[�V�����t���O
Const ANIMATION_FLAG = False
'�ŏI�t���[No�i�萔�����Ȃ̂ő啶����`�ɂ���j
Dim LAST_FLOW_NO As String
'�������W
Dim CENTER_POINT As Integer
'����̒������W
Dim LEFT_POINT As Integer
'�V�F�C�v���X�g
Dim shapeList As Variant
'�t���[���X�g
Dim flowList As Variant
'�쐬�V�F�C�v
Dim onShape As shape
'Y���W
Dim yPoint As Integer
'�ړ��ς݃V�F�C�v�ԍ�
Dim movedShapeNos As String
'�W���V�F�C�v
'Dim baseShape As Shape
Sub A_�t���[�쐬()
Attribute A_�t���[�쐬.VB_ProcData.VB_Invoke_Func = "q\n14"

    '��1.�����ݒ�
    Call init

    '��2.�V�F�C�v�쐬
    Call createFlowParts

    '��3.�ꏊ�ړ�
    Call moveFlowParts

    '��4.�R�l�N�^�t�^
    Call addConnector

    '��5.����
    Call adjust

End Sub
'��1.�����ݒ�
Function init()

    '�������W
     CENTER_POINT = Selection.Left + Selection.Width / 2
     LEFT_POINT = CENTER_POINT - 200

    '�V�F�C�v��ʒl�iAutoShapeType�j���œK��
    With Sheets("�V�F�C�v�ꗗ")
    
        Dim lastShapeLine As Integer: lastShapeLine = .Cells(1, 1).End(xlDown).Row
        For i = 2 To lastShapeLine
            For Each shp In .Shapes
                If .Cells(i, 6).Left < shp.Left _
                    And .Cells(i, 6).Top < shp.Top _
                    And .Cells(i, 6).Top + .Cells(i, 6).Height > shp.Top + shp.Height Then

                    .Cells(i, 3) = shp.AutoShapeType

                    Exit For
                End If
            Next
        Next

        '�V�F�C�v���X�g
        shapeList = .Range(.Cells(2, 1), .Cells(lastShapeLine, 5))
    End With
    
    With Sheets("�ڍ�")

        Dim lastFlowLine As Integer: lastFlowLine = .Cells(1, 2).End(xlDown).Row

        '�t���[�̍ŏI�ԍ�
        LAST_FLOW_NO = .Cells(lastFlowLine, 2)

        '�t���[���X�g�i�݌v���j
        flowList = .Range(.Cells(2, 1), .Cells(lastFlowLine, 5))
    
    End With
    yPoint = Selection.Top
End Function

'��2.�V�F�C�v�쐬
Function createFlowParts()

    For i = LBound(flowList) To UBound(flowList)
        '�V�F�C�v��ʎ擾
        shapeType = vlookup(shapeList, flowList(i, TYPE_COL), 2, 3)
        '�V�F�C�v�𐶐�
        Set onShape = ActiveSheet.Shapes.AddShape(shapeType, 40, 10, 100, 30)
        '�e�L�X�g�̐ݒ�
        onShape.TextFrame.Characters.text = flowList(i, PROCESS_NO_COL) & "." & flowList(i, SHAPE_TEXT_COL)
        '���O�̐ݒ�i�t���[No�j
        onShape.Name = flowList(i, FLOW_NO_COL)
        '�V�F�C�v���œK��
        Set onShape = baseShape(onShape)

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
'��3.�ꏊ�ړ�
Function moveFlowParts()
    '�ړ�������V�F�C�v
    Dim moveShape As shape
    Dim srcShapes As Variant
    Dim beforeCenterShape As String
    '�J�n�v�f�̈ړ�
    Call moveStartEnd(ActiveSheet.Shapes("1"))
    
    For i = LBound(flowList) + 1 To UBound(flowList) - 1
        Set moveShape = ActiveSheet.Shapes(CStr(flowList(i, FLOW_NO_COL)))
        '�J�ڌ��V�F�C�v���z����擾
        srcShapes = getSrc(moveShape, CInt(i))
        '�J�ڌ��V�F�C�v�̂����A�������W�̂��̂��擾
        beforeCenterShape = getMainPointShape(srcShapes)
        
        If moveShape.Name = "5" Then
            Debug.Print ""
        End If
        '���m���̒Ⴂ����قǑO�ɂ����Ă���̂������񂩂�
        
        
        If UBound(srcShapes) = 0 And isSwitchShape(srcShapes, moveShape) <> "" Then '�����̒��ł�������B�֐������ăX�}�[�g�ɁB
            '�J�ڌ���1�݂̂��¤���ꂪSwitch�v�f
            Dim currentNo As Integer: currentNo = isSwitchShape(srcShapes, moveShape)
            If currentNo = 1 Then
                '�����V�F�C�v���ŏ��̗v�f�̂Ƃ��̂݃g�b�v��ݒ�
                yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
            End If
            Dim branchCount As Integer: branchCount = getSwitchBranchCount(srcShapes) + 1
            '�ړ��̏d��
            Dim weight As Integer
            If branchCount Mod 2 = 0 Then
                '����
                Call switchEvenChildlenMove(moveShape, branchCount, currentNo)
            Else
                '�
                Call switchOddChildlenMove(moveShape, branchCount, currentNo)
            End If
            
'            moveShape.top = yPoint
            Call animationTop(moveShape, yPoint)
            
        ElseIf UBound(srcShapes) = 0 And isBranchShape(srcShapes, moveShape) Then
            '�J�ڌ����P�݂̂��A���J�ڌ�������ł��̑�ꕪ��悪���ؒ��V�F�C�v�i���ړ�����j
            'yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN '�ǂ����ł��������A�R�l�N�^�����ɂ������
'            moveShape.top = yPoint
            Call animationTop(moveShape, yPoint)
            Dim srtShapePoint As Integer: srtShapePoint = getCenterPoint(ActiveSheet.Shapes(srcShapes(0)))
'            moveShape.Left = srtShapePoint - moveShape.Width / 2 - 170 '�J�ڌ��V�F�C�v�Ƃ̍����w�肷��
            Call animationLeft(moveShape, srtShapePoint - moveShape.Width / 2 - 170)
        
        ElseIf UBound(srcShapes) = 0 And Not isCenterShape(srcShapes) Then
            '�J�ڌ����P�݂̂��A���ꂪ�����ł͂Ȃ�
            yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
'            moveShape.top = yPoint
'            moveShape.Left = CENTER_POINT - moveShape.Width / 2 - 170
            Call animationTop(moveShape, yPoint)
            Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 - 170)
            
        ElseIf UBound(srcShapes) = 0 And isCenterShape(srcShapes) Then
            '�J�ڌ����P�݂̂��A���ꂪ����
            yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
            
            Call animationTop(moveShape, yPoint)
            Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2)
            
        ElseIf UBound(srcShapes) > 0 Then  'And isCenterShape(srcShapes)
            '�J�ڌ���2�ȏ゠��A���̒��Ƀ��C�����W�̂��̂�����i���̏����������j
            yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
'            moveShape.top = yPoint
'            moveShape.Left = CENTER_POINT - moveShape.Width / 2
            Call animationTop(moveShape, yPoint)
            Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2)
        Else
        End If
    Next
    
    '�I���v�f�̈ړ�
    Call moveStartEnd(ActiveSheet.Shapes(LAST_FLOW_NO))
End Function
Function switchEvenChildlenMove(moveShape As shape, branchCount As Integer, currentNo As Integer)
    Dim weight As Integer
    If branchCount / 2 >= currentNo Then
        '�O��
        weight = branchCount / 2 - currentNo + 1
'       moveShape.Left = CENTER_POINT - moveShape.Width / 2 - (140 * weight) + 70
        Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 - (140 * weight) + 70)
    Else
        '�㔼
        weight = currentNo - branchCount / 2
'       moveShape.Left = CENTER_POINT - moveShape.Width / 2 + (140 * weight) - 70
        Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 + (140 * weight) - 70)
    End If
End Function
Function switchOddChildlenMove(moveShape As shape, branchCount As Integer, currentNo As Integer)
    If branchCount / 2 + 0.5 >= currentNo Then
        '�O��+����
        weight = branchCount / 2 + 0.5 - currentNo
'       moveShape.Left = CENTER_POINT - moveShape.Width / 2 - (140 * weight)
        Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 - (140 * weight))
    Else
        '�㔼
        weight = currentNo - (branchCount / 2 + 0.5)
'       moveShape.Left = CENTER_POINT - moveShape.Width / 2 + (140 * weight)
        Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2 + (140 * weight))
    End If

End Function
Function animationTop(moveShape As shape, goalPoint As Integer)
    If ANIMATION_FLAG Then
        While Not isApproximate(moveShape.Top, goalPoint, 2)
            If isApproximate(moveShape.Top, goalPoint, 10) Then
                moveShape.Top = moveShape.Top + 1
            Else
                moveShape.Top = moveShape.Top + 9
            End If
            Application.wait [Now() + "0:00:00.0005"]
        Wend
    Else
        moveShape.Top = goalPoint
    End If
End Function
Function animationLeft(moveShape As shape, goalPoint As Integer)
    If ANIMATION_FLAG Then
        While moveShape.Left <> goalPoint
            If isApproximate(moveShape.Left, goalPoint, 10) Then
                moveShape.Left = moveShape.Left + 1
            Else
                moveShape.Left = moveShape.Left + 9
            End If
            Application.wait [Now() + "0:00:00.0005"]
        Wend
    Else
        moveShape.Left = goalPoint
    End If
End Function
Function isCenterShape(srcShapes As Variant)
    For Each no In srcShapes
        '�ߎ��l�`�F�b�N�ɂ�����
        If isApproximate(getCenterPoint(ActiveSheet.Shapes(no)), CENTER_POINT) Then
            isCenterShape = True
            Exit Function
        End If
    Next
End Function
Function getCenterPoint(argShp As shape)
    getCenterPoint = argShp.Left + argShp.Width / 2
End Function
'Switch�v�f�̕��򐔂�Ԃ��B�V�F�C�v���X�g�̂����A�ŏ��Ɍ�������Switch�v�f���ΏۂɂȂ�
Function getSwitchBranchCount(srcShapes As Variant)
    For Each no In srcShapes
        distShape = vlookup(flowList, no, 2, 3)
        flowType = vlookup(flowList, no, 2, 4)
        distShapeArray = Split(distShape, vbLf)
        
        If flowType = "Switch" Then
            getSwitchBranchCount = UBound(distShapeArray)
        End If
    Next
End Function

Function isSwitchShape(srcShapes As Variant, moveShape As shape)
    Dim flowType As String
    Dim distShape As String
    Dim distShapeArray As Variant
    For Each no In srcShapes
        distShape = vlookup(flowList, no, 2, 3)
        flowType = vlookup(flowList, no, 2, 4)
        distShapeArray = Split(distShape, vbLf)
        
        If flowType = "Switch" Then
            For i = 0 To UBound(distShapeArray)
                If Split(distShapeArray(i), ":")(1) = moveShape.Name Then
                    isSwitchShape = i + 1
                    Exit Function
                End If
            Next
        End If
    Next
End Function
'�w�肳�ꂽ�V�F�C�v�z��ɕ���v�f�����邩�A�܂����̑�ꕪ���͍����ؒ��̃V�F�C�v���ǂ���
Function isBranchShape(srcShapes As Variant, moveShape As shape)
    Dim flowType As String
    Dim distShape As String
    Dim distShapeArray As Variant
    For Each no In srcShapes
        '����
        distShape = vlookup(flowList, no, 2, 3)
        flowType = vlookup(flowList, no, 2, 4)
        distShapeArray = Split(distShape, vbLf)
        If flowType = "����" Then
            If Split(distShapeArray(0), ":")(1) = moveShape.Name Then
                isBranchShape = True
                Exit Function
            End If
        End If
    Next
End Function
'�w�肳�ꂽ�J�ڌ��V�F�C�v�̂����A���C�����W�ɂ�����̂̃V�F�C�v����Ԃ�
Function getMainPointShape(srcShapes As Variant)
    For Each no In srcShapes
        '�ߎ��l�`�F�b�N
        If isApproximate(CENTER_POINT, ActiveSheet.Shapes(no).Left + ActiveSheet.Shapes(no).Width / 2) Then
            mainStreetFlag = True
            getMainPointShape = no
            Exit Function
        End If
    Next
End Function
'�J�ڌ��̃V�F�C�v����
Function getSrc(moveShape As shape, currrentNo As Integer)
    Dim srcNos As String
    For j = LBound(flowList) To currrentNo
        If moveShape.Name = CStr(flowList(j, DIST_NO_COL)) Then
            '�P�ꔭ���B�J�ڌ��i�[
            srcNos = srcNos & CStr(flowList(j, FLOW_NO_COL)) & ","
    
        ElseIf CStr(flowList(j, DIST_NO_COL)) Like "*:" & moveShape.Name & vbLf & "*" _
            Or CStr(flowList(j, DIST_NO_COL)) Like "*:" & moveShape.Name Then
            '2���̌댟�m��h�~���邽�߁A���̉��s�����A�O���݈̂�v�������Ƃ���
            '���������B�J�ڌ��i�[
            srcNos = srcNos & CStr(flowList(j, FLOW_NO_COL)) & ","
        End If
    Next
    '�z��ŕԂ�
    getSrc = Split(deleteEndText(srcNos), ",")
End Function
Function deleteEndText(text As String, Optional deleteLength As Long = 1) As String
    If Len(text) >= deleteLength Then
        deleteEndText = Left(text, Len(text) - deleteLength)
    Else
        deleteEndText = text
    End If
End Function
'�J�n�A�I���v�f�ɑ΂���V�F�C�v�ړ����{��
Function moveStartEnd(moveShape As shape)
    moveShape.Width = 75
'    moveShape.Left = CENTER_POINT - moveShape.Width / 2
            Call animationLeft(moveShape, CENTER_POINT - moveShape.Width / 2)
    yPoint = yPoint + moveShape.Height + HEIGHT_MARGIN
'    moveShape.top = yPoint
            Call animationTop(moveShape, yPoint)
    movedShapeNos = movedShapeNos & moveShape.Name & ","
End Function
'
Function wait(waitTime As String)
    If ANIMATION_FLAG Then
        Application.wait [Now() + waitTime]
    End If
End Function
'��4.�R�l�N�^�t�^
Function addConnector()
    Dim srcShape As shape
    Dim distShape As shape
    Dim connectShape As shape
    Dim startPoint As Integer
    Dim endPoint As Integer
    Dim hjShape As shape

    For i = LBound(flowList) To UBound(flowList) - 1
        Set srcShape = ActiveSheet.Shapes(CStr(flowList(i, FLOW_NO_COL)))
        startPoint = vlookup(shapeList, srcShape.AutoShapeType, 3, 4)
        
        If CStr(flowList(i, DIST_NO_COL)) Like "*" & vbLf & "*" Then
            '�J�ڐ敡��
            Dim distArray As Variant: distArray = Split(CStr(flowList(i, DIST_NO_COL)), vbLf)
            For j = LBound(distArray) To UBound(distArray)
                Set connectShape = baseConnect(connectShape)
                connectShape.ConnectorFormat.BeginConnect srcShape, startPoint
                
            If ANIMATION_FLAG Then
                Application.wait [Now() + "0:00:00.1"]
            End If

                Set distShape = ActiveSheet.Shapes(Split(CStr(distArray(j)), ":")(1))
                endPoint = vlookup(shapeList, distShape.AutoShapeType, 3, 5)
                connectShape.ConnectorFormat.EndConnect distShape, endPoint
                
                '�������W���݂ăR�l�N�^��ނ�ύX
                If Not isApproximate(srcShape.Left + srcShape.Width / 2, distShape.Left + distShape.Width / 2) Then
                    connectShape.ConnectorFormat.Type = msoConnectorElbow
                End If
                If isApproximate(srcShape.Top + srcShape.Height / 2, distShape.Top + distShape.Height / 2) Then
                    connectShape.ConnectorFormat.Type = msoConnectorStraight
                End If
                
                If UBound(distArray) = 1 And j = 0 Then
                    '�����1�{��
                    connectShape.ConnectorFormat.BeginConnect srcShape, 2
                    '�ŏI�|�C���g�͉E���ɂ��邩��ڑ��|�C���g���̎w��Ƃ���
                    connectShape.ConnectorFormat.EndConnect distShape, distShape.ConnectionSiteCount
                End If
                connectShape.Name = srcShape.Name & "-" & distShape.Name
                '�ē��V�F�C�v
                Set hjShape = ActiveSheet.Shapes.AddShape(61, 40, 10, 10, 10)
                hjShape.TextFrame.Characters.text = Split(CStr(distArray(j)), ":")(0)
                hjShape.Name = connectShape.Name & "support"
                Set hjShape = supportShape(hjShape)
                If UBound(distArray) = 1 Then
                    '����
                    hjShape.Left = connectShape.Left + 2
                    hjShape.Top = connectShape.Top + 2
                Else
                    'Switch
                    hjShape.Left = distShape.Left + 63
                    hjShape.Top = distShape.Top - 16
                End If
            If ANIMATION_FLAG Then
                Application.wait [Now() + "0:00:00.1"]
            End If
            Next
        Else
            '�J�ڐ�P��
            Set distShape = ActiveSheet.Shapes(CStr(flowList(i, DIST_NO_COL)))
            Set connectShape = baseConnect(connectShape)
            
            connectShape.ConnectorFormat.BeginConnect srcShape, startPoint
            If ANIMATION_FLAG Then
                Application.wait [Now() + "0:00:00.1"]
            End If
            endPoint = vlookup(shapeList, distShape.AutoShapeType, 3, 5)
            connectShape.ConnectorFormat.EndConnect distShape, endPoint
            connectShape.Name = srcShape.Name & "-" & distShape.Name
            
            '�������W���݂ăR�l�N�^��ނ�ύX
            If Not isApproximate(srcShape.Left + srcShape.Width / 2, distShape.Left + distShape.Width / 2) Then
                connectShape.ConnectorFormat.Type = msoConnectorElbow
            End If
            
            '�J�ڐ悪�u�I���v�t���[��������n�_�I�_��ς��� (Not isApproximate(srcShape.Left + srcShape.Width / 2, distShape.Left + distShape.Width / 2)) And
            If distShape.Name = LAST_FLOW_NO Then
                If distShape.Top - srcShape.Top < 100 Then
                    connectShape.ConnectorFormat.BeginConnect srcShape, srcShape.ConnectionSiteCount - 1
                    connectShape.ConnectorFormat.EndConnect distShape, 1
                Else
                    '�I���V�F�C�v���牓���ꍇ
                    connectShape.ConnectorFormat.BeginConnect srcShape, srcShape.ConnectionSiteCount - 2
                    connectShape.ConnectorFormat.EndConnect distShape, 2
                End If
            End If
            If ANIMATION_FLAG Then
                Application.wait [Now() + "0:00:00.1"]
            End If
        End If
    Next

End Function
'��5.����
Function adjust()
    '�S�V�F�C�v���݂Đ��̏d�����Ȃ������`�F�b�N������
    '�������s���Ƃ��Ȃ�
'    Dim beforeLeft As Integer
'    Dim leftest As Integer
'    Dim leftestShape As shape
'    For Each shp In ActiveSheet.Shapes
'        If shp.Connector Then
'
'            If shp.Left < beforeLeft Then
'                leftest = shp.Left
'                Set leftestShape = shp
'            End If
'            beforeLeft = shp.Left
'        End If
'    Next
End Function
'�ߎ��l���ǂ���
Function isApproximate(int1 As Integer, int2 As Integer, Optional diff As Integer = 1)
    If int1 < int2 Then
        isApproximate = int2 - int1 <= diff
    Else
        isApproximate = int1 - int2 <= diff
    End If
End Function
'�V�F�C�v�œK��
Function baseShape(onShape As shape)
    '���h��Ԃ��imsoTrue:����AmsoFalse:�Ȃ��j
    onShape.Fill.Visible = msoTrue
    '�����̑����B���D�݂łǂ���
    onShape.Line.weight = 1
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
    'onShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    '�Z���ɍ��킹�Ĉړ���T�C�Y�ύX�����Ȃ�
    onShape.Placement = xlFreeFloating
    '����
    onShape.Height = SHAPE_HEIGHT
    '��
    onShape.Width = SHAPE_WIDTH

    If onShape.Width < SHAPE_WIDTH Then
        onShape.TextFrame2.AutoSize = msoAutoSizeNone
    Else
        'onShape.Width = SHAPE_HEIGHT + 20
        'onShape.Width = onShape.Width + 20
    End If

    onShape.TextFrame2.TextRange.Font.NameComplexScript = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.NameFarEast = "Meiryo UI"
    onShape.TextFrame2.TextRange.Font.Name = "Meiryo UI"
    onShape.TextFrame2.WordWrap = msoTrue
    onShape.TextFrame2.AutoSize = msoAutoSizeNone
    onShape.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
    'TODO:�Z�����c�ɍL�����Ă��A�ς��Ȃ��I�v�V���������邱��
    Set baseShape = onShape
End Function
'�ē��V�F�C�v
Function supportShape(onShape As shape)
    '���h��Ԃ��imsoTrue:����AmsoFalse:�Ȃ��j
    onShape.Fill.Visible = msoFalse
    '�����̑����B���D�݂łǂ���
    onShape.Line.weight = 0
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
    Set supportShape = onShape
End Function
'�R�l�N�^
Function baseConnect(connectShape As shape)
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
    connectShape.Line.weight = 1
    Set baseConnect = connectShape
End Function
'�\����g�p�������V�F�C�v��T��
Function vlookup(list As Variant, searchVal As Variant, searchCol As Integer, returnCol As Integer)
    '�����l
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
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).Top + Selection.ShapeRange.Item(1).Height / 2
    For Each sp In Selection.ShapeRange
        sp.Top = baseXCenterPoint - sp.Height / 2
    Next
End Sub
'�I�������R�l�N�^�̎n�_���̃V�F�C�v�Ƃ̐ڑ��ʒu��ύX����
Sub E_�R�l�N�^�n�_�ύX()
Attribute E_�R�l�N�^�n�_�ύX.VB_ProcData.VB_Invoke_Func = "r\n14"

    If connectorErrorCheck Then
       Exit Sub
    End If

    Dim currentBeginConnectPoint As Integer
    Dim targetShape As shape

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
    Dim targetShape As shape

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
