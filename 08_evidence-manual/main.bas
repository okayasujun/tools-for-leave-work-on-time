Attribute VB_Name = "main"
'###############################
'�@�\���F�G�r�f���X�E�}�j���A���쐬�x���c�[�� v2.0
'Author�Fokayasu jun
'�쐬���F2022/10/19
'�X�V���F2023/05/13
'COMMENT�F�e�R�����g�́u���v�͕ύX�\�������B�p�r��D�݂ɍ��킹�ĕς��Ă݂āB
'###############################
'�|�C���^API�B�}�E�X�J�[�\���ʒu����Z���ʒu���擾���邽�߂Ɏg�p����
Private Type POINTAPI
    x As Long
    y As Long
End Type
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'���摜�L���v�V�����ɕ�L������ꍇ�̍s���i�摜�ƃL���v�V�����̊Ԃ̋�s���j
Const REMARK_LINE = 0
Const CAPTION_TEXT_TOP_FLAG = True
'#�I���Z���͈͂̑傫���̐Ԙg���}�E�X�ʒu�ɏo��������
Sub AA_�Ԙg���o��������()
Attribute AA_�Ԙg���o��������.VB_ProcData.VB_Invoke_Func = "q\n14"

    '�Ԙg�V�F�C�v���i�[����ϐ�
    Dim onShape As Shape
    '�����J�n���ɑI�����Ă���Z�����
    Dim beforeRange As Range
    
    On Error GoTo ErrHndl
    '�����O��őI���Z���ʒu��ێ����邽��
    Set beforeRange = Selection
    
    '�V�F�C�v����&�X�^�C���ݒ�
    '���������͈ȉ���URL���Q�l�ɕύX�B�V�F�C�v�̌`���w�肷��
    'https://learn.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype?redirectedfrom=MSDN
    Set onShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, _
                                            Selection.left, _
                                            Selection.top, _
                                            Selection.WIDTH, _
                                            Selection.Height)
    '���h��Ԃ��imsoTrue:����AmsoFalse:�Ȃ��j
    onShape.Fill.Visible = msoFalse
    '�����̑���
    onShape.Line.Weight = 4
    '���F�w��
    onShape.Line.ForeColor.RGB = RGB(255, 0, 0)
    '���h��Ԃ��F�w��
    'onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    '�����̃X�^�C���i����/�_���j
    'onShape.Line.DashStyle = msoLineDash
    '�����̃X�^�C���i��d��/��d���j
    'onShape.Line.Style = msoLineThinThin
    
    '���I���Z���ʒu�ɕ\�����邾���ł����ꍇ�͂����ŏ������I������B�R�����g�C������
    'Exit Sub
    
    '�摜�̏�ɃJ�[�\���������Ԃł͌㑱�������ł��Ȃ����߁A�ꎞ��\���ɂ���
    '���̒i�K�łǂ̃V�F�C�v��ɃJ�[�\�������邩�s���Ȃ̂őS�V�F�C�v��ΏۂƂ���
    For Each shp In ActiveSheet.Shapes
        shp.Visible = False
    Next
    
    '���W�擾�����̂��߂ɁA�}�E�X�J�[�\���̏�Ԃ��������
    beforeRange.Select
    
    '�ȉ��ŏ�L�쐬�V�F�C�v���}�E�X������Z���̈ʒu�Ɉړ�������
    Dim p As POINTAPI
    Dim Getcell As Range

    '�J�[�\���ʒu���擾
    GetCursorPos p

    '�}�E�X�J�[�\���̈ʒu����Z�����擾
    Set Getcell = ActiveWindow.RangeFromPoint(p.x, p.y)

    '�V�F�C�v�ʒu���}�E�X�J�[�\���̒��߃Z���̍���ɍ��킹��
    onShape.top = Getcell.top
    onShape.left = Getcell.left
    
    '�S�V�F�C�v������Ԃɖ߂�
    For Each shp In ActiveSheet.Shapes
        shp.Visible = True
    Next
    Exit Sub
ErrHndl:
    'MsgBox "�G���[����"
    '�S�V�F�C�v������Ԃɖ߂�
    For Each shp In ActiveSheet.Shapes
        shp.Visible = True
    Next

End Sub
'#�R�l�N�^�n�ȊO�̃V�F�C�v�y�щ摜�ɉe������i�h��Ԃ��Ȃ��̃I�[�g�V�F�C�v�͑ΏۊO�j
'�O���[�v�͂��̍\���V�F�C�v���ׂĂɉe�����Ă��܂���������ɂ��Ȃ�
Sub AB_�e��t����()
Attribute AB_�e��t����.VB_ProcData.VB_Invoke_Func = "w\n14"

    Dim shp As Shape

    If TypeName(Selection) = "Range" Then
        '�V�F�C�v���I����ԁB�S�V�F�C�v��Ώۂɂ���
        For Each shp In ActiveSheet.Shapes
            '�����ύX�������Ƃ��͂�������Q�l�ɁFhttps://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
            If shp.Type = msoPicture Or (shp.Type = msoAutoShape And shp.Fill.Visible) Then
                Call castShadow(shp)
            End If
        Next
    Else
        '�I�𒆃V�F�C�v�̂ݏ������s��
        For Each shp In Selection.ShapeRange
            If shp.Type = msoPicture Or (shp.Type = msoAutoShape And shp.Fill.Visible) Then
                Call castShadow(shp)
            End If
        Next
    End If
End Sub
'�w�肳�ꂽ�V�F�C�v�ɉe��t����
Function castShadow(shp As Shape)
    With shp.Shadow
        '���e�̎�ށFhttps://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.core.msoshadowtype?view=office-pia
        .Type = msoShadow26
        '�e�̕\���ؑ�
        .Visible = msoTrue
        '���e�̌��ʁFhttps://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.core.msoshadowstyle?view=office-pia
        .Style = msoShadowStyleOuterShadow
        '���u���A�[�B�e�̂ڂ����
        .Blur = 20
        '���e�̑��Έʒu
        .OffsetX = 7.7781745931
        .OffsetY = 7.7781745931
        '���e���V�F�C�v�ƂƂ��ɉ�]�����邩�ǂ���
        .RotateWithShape = msoFalse
        '���e�̐F
        .ForeColor.RGB = RGB(100, 100, 100)
        '���g�����X�p�����V�[�B�e�̓����x�B0�`1�Ŏw��B1�͊��S�ɓ���
        .Transparency = 0.4
        '�e�̃T�C�Y
        .Size = 100
    End With
End Function
'#�摜�̌��ʂ����Z�b�g����
Sub AC_���ʂ����Z�b�g����()
Attribute AC_���ʂ����Z�b�g����.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim shp As Shape
    If TypeName(Selection) = "Range" Then
        '�V�F�C�v���I����ԁB�S�V�F�C�v��Ώۂɂ���
        For Each shp In ActiveSheet.Shapes
            Call shapeReset(shp)
        Next
    Else
        '�I�𒆃V�F�C�v�̂ݏ������s��
        For Each shp In Selection.ShapeRange
            Call shapeReset(shp)
        Next
    End If
End Sub
'�w�肳�ꂽ�V�F�C�v�y�щ摜�̌��ʂ����������܂��B
'�R�l�N�^�n�ȊO�̃V�F�C�v�y�щ摜�����Z�b�g����i�h��Ԃ��Ȃ��̃I�[�g�V�F�C�v�͑ΏۊO�j
Function shapeReset(shp As Shape)
    If shp.Type = msoPicture Or (shp.Type = msoAutoShape And shp.Fill.Visible) Then
        With Application.CommandBars
            '�u�}�̃��Z�b�g�v���\�ȂƂ��̂�
            If .GetEnabledMso("PictureReset") Then
                .ExecuteMso "PictureReset"
            End If
            '�I�[�g�V�F�C�v�̏ꍇ�͈ȉ��ŊO���i����肽�����Ƃɉ����ď������e��ǉ����邱�Ɓj
            shp.Shadow.Visible = msoFalse
        End With
    End If
End Function
'#�V�F�C�v�𐮗񂳂���i�\�t���j
Sub AD_�V�F�C�v��\�t���ɐ��񂳂���()
Attribute AD_�V�F�C�v��\�t���ɐ��񂳂���.VB_ProcData.VB_Invoke_Func = "e\n14"
    '���摜�Ԃ̊Ԋu
    Const MARGIN_BOTTOM = 70
    
    '�\�t���W���i�[����itop�͓s�x���������Aleft�͏����l���g���܂킷�j
    Dim top As Integer: top = Selection.top + 5
    
    '�L���v�V�������L�ڂ���p�̃Z��
    Dim captionRange As Range
    Dim moveShape As Shape
    
    '�G���[�`�F�b�N
    If Selection.Row - REMARK_LINE - 1 < 1 Then
        MsgBox "�L���v�V�����p�̍s������܂���B����" & REMARK_LINE - Selection.Row + 2 & "�s���̈ʒu�Ŏ��s���Ă��������B"
        Exit Sub
    End If
    
    '�L���v�V�����^�C�g��
    Dim captionText As String
    '���_�C�A���O���g���ꍇ�͈ȉ��̃R�����g�A�E�g�������g�p����
    If CAPTION_TEXT_TOP_FLAG Then
        captionText = "��" 'InputBox("�L���v�V�����̏����l�����āB", "�L���v�V�����I�v�V����", "�������ɉ摜�̐���������")
    Else
        captionText = "��" 'InputBox("�L���v�V�����̏����l�����āB", "�L���v�V�����I�v�V����", "�������ɉ摜�̐���������")
    End If
    
    If StrPtr(captionText) = 0 Then
        '�L�����Z����
        Exit Sub
    End If
    
    For Each moveShape In ActiveSheet.Shapes
        '���ɊY��������̂��ΏہF�摜�A�O���[�v�i���K�v�ɉ����ď����������āj
        '�����ύX���Q�l�Fhttps://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
        If moveShape.Type <> msoPicture _
            And moveShape.Type <> msoGroup _
            And (moveShape.Type = msoAutoShape) Then 'And Not moveShape.Fill.Visible
            GoTo CONTINUE:
        End If
        
        '�V�F�C�v���ړ������āA
        Set captionRange = move(moveShape, top)
        
        '���L���v�V�������͂̐ݒ�i�s�v�Ȃ�R�����g�A�E�g���āj
        Call setCaption(captionRange, captionText)
        
        '���Ώۂɂ����V�F�C�v�̏㕔���W + ���Ώۂɂ����V�F�C�v�̍��� + �摜�Ԃ̊Ԋu + �L���v�V�����Z���s�̍��� = ���̃V�F�C�v�̈ړ���㕔���W
        top = top + moveShape.Height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).Height
CONTINUE:
    Next
    
    'END����
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
    Call setCaption(dummyShape.TopLeftCell, "- END -")
    dummyShape.Delete
    
End Sub
'�w�肳�ꂽ�V�F�C�v��������̈ʒu�Ɉړ�������
Function move(moveShape As Shape, top As Integer)
    '�ړ��ʒu���擾���邽�߂̃_�~�[�V�F�C�v
    Dim dummyShape As Shape
    Dim left As Integer: left = Selection.left
    
    '������̃Z�����擾���邽�߂̃_�~�[�V�F�C�v
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
        
    '�V�F�C�v���ړ�����
    moveShape.top = dummyShape.TopLeftCell.Offset(0, 0).top
    moveShape.left = Selection.left
    
    '��ƋL�^�̈Ӗ������ŉ摜���Ɏ��Ԃ�ݒ肷��B��ӂɂ��邽�߃~���b�������ɕt�^
    moveShape.Name = "image-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
    
    If CAPTION_TEXT_TOP_FLAG Then
        '�L���v�V�������͗p�Z�����擾����i-1�̓^�C�g�����j
        Set move = dummyShape.TopLeftCell.Offset(-1 - REMARK_LINE, 0)
    Else
        dummyShape.top = moveShape.top + moveShape.Height
        '�L���v�V�������͗p�Z�����擾����i-1�̓^�C�g�����j
        Set move = dummyShape.TopLeftCell.Offset(1, 0)
    End If
    
    '�p�ς݂�����폜����
    dummyShape.Delete
End Function
'�L���v�V�����p�Z���̐ݒ�
Function setCaption(captionRange As Range, captionText As String)
    '�摜�Ԉړ���Ctrl+���ō����ɍs������
    captionText = IIf(captionText = "", " ", captionText)
    '���K�X�ς��Ă悵�B���D�݂�
    captionRange.Value = captionText
    captionRange.Font.Bold = True
    captionRange.Font.Color = RGB(40, 40, 40)
    captionRange.Font.Size = 18
    captionRange.Font.Name = "BIZ UDP�S�V�b�N" '"Meiryo UI"'
End Function
'�~���b���擾
Function getMSec() As String
    Dim dblTimer As Double
    Dim s_return As String
    dblTimer = CDbl(Timer)
    s_return = Format(Fix((dblTimer - Fix(dblTimer)) * 1000), "000")
    getMSec = s_return
End Function
'�I�𒆂̃V�F�C�v��I�����ɃR�l�N�^�Ōq��
Sub AF_�V�F�C�v��I�����ɃR�l�N�^�Ōq��()
Attribute AF_�V�F�C�v��I�����ɃR�l�N�^�Ōq��.VB_ProcData.VB_Invoke_Func = "l\n14"
    Dim startShape As Shape
    Dim endShape As Shape
    Dim connectShape As Shape
    
    If TypeName(Selection) = "Range" Then
        MsgBox "�V�F�C�v���I������Ă��܂���B2�ȏ�I�����Ă��������B"
        Exit Sub
    End If
    For Each shp In Selection.ShapeRange
        If shp.Type = msoGroup Or shp.Connector Then
            MsgBox "�I���V�F�C�v�ɃO���[�v���R�l�N�^���܂܂�Ă��܂��B�������Ă��������B"
            Exit Sub
        End If
    Next
    
    For i = 1 To Selection.ShapeRange.count - 1
        '�I�𒆃V�F�C�v�̕ێ��i�ڑ����j
        Set startShape = Selection.ShapeRange.Item(i)
        '�I�𒆃V�F�C�v�̕ێ��i�ڑ���j
        Set endShape = Selection.ShapeRange.Item(i + 1)

        '�ڑ��V�F�C�v�̒a��
        '��Type�����͉E�L���Q�ƁFhttps://learn.microsoft.com/ja-jp/office/vba/api/office.msoconnectortype
        Set connectShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorStraight, BeginX:=0, BeginY:=0, EndX:=0, EndY:=0)
        '���ڑ��̎n�_�ʒu�w��i�Ō�̈�����1:��ӁA2:���ӁA3:���ӁA4�E�Ӂj
        connectShape.ConnectorFormat.BeginConnect ActiveSheet.Shapes(startShape.Name), 4
        '���ڑ��̏I�_�ʒu�w��i�Ō�̈����͎n�_�ʒu�̎w��Ɠ��l�j
        connectShape.ConnectorFormat.EndConnect ActiveSheet.Shapes(endShape.Name), 2
        '�R�l�N�^�����H����
        Call makeConnectAllow(connectShape)
    Next
End Sub
'�R�l�N�^�����H����B�e�v���p�e�B�D�݂ɍ��킹�Đݒ肳�ꂽ��
Function makeConnectAllow(connectShape As Shape)
    '���I�_�R�l�N�^���O�p�ɁB
    connectShape.Line.EndArrowheadStyle = msoArrowheadTriangle
    '�����̐F
    connectShape.Line.ForeColor.RGB = RGB(10, 10, 10)
    '�����̑���
    connectShape.Line.Weight = 1
    '���I�_�̒���
    connectShape.Line.EndArrowheadLength = msoArrowheadLong
    '���I�_�̑���
    connectShape.Line.EndArrowheadWidth = msoArrowheadWide
    '���O
    connectShape.Name = "connect-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
End Function
'�R�l�N�^�̎�ނ𒼐��A�Ȑ��A�G���{�[�Ő؂�ւ���
Sub AG_�R�l�N�^��ސ؂�ւ�()
Attribute AG_�R�l�N�^��ސ؂�ւ�.VB_ProcData.VB_Invoke_Func = "i\n14"

    If connectorErrorCheck Then
       Exit Sub
    End If
    
    With Selection.ShapeRange.ConnectorFormat
        If .Type = msoConnectorElbow Then
            .Type = msoConnectorCurve
        ElseIf .Type = msoConnectorCurve Then
            .Type = msoConnectorStraight
        Else
            .Type = msoConnectorElbow
        End If
    End With
End Sub
'�I�������R�l�N�^�̎n�_���̃V�F�C�v�Ƃ̐ڑ��ʒu��ύX����
Sub AH_�R�l�N�^�n�_�ύX()
Attribute AH_�R�l�N�^�n�_�ύX.VB_ProcData.VB_Invoke_Func = "o\n14"

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
Sub AI_�R�l�N�^�I�_�ύX()
Attribute AI_�R�l�N�^�I�_�ύX.VB_ProcData.VB_Invoke_Func = "p\n14"
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
'#�I�𒆂̃V�F�C�v���O���[�v��
Sub AJ_�I�𒆃V�F�C�v���O���[�v��()
    Selection.Group.Select
End Sub
'#�I�𒆂̃V�F�C�v���O���[�v����
Sub AK_�I�𒆃V�F�C�v���O���[�v����()
    Selection.Ungroup
End Sub
'#�I�𒆂̃V�F�C�v���Ŕw�ʂɂ���
Sub AL_�I�𒆃V�F�C�v���Ŕw�ʂ�()
    If TypeName(Selection) = "Range" Then
        MsgBox "�V�F�C�v����I��ł�����s���ĂˁB"
        Exit Sub
    End If
    
    For Each shp In Selection.ShapeRange
        shp.ZOrder msoSendToBack
    Next
End Sub
'#�I�𒆂̃V�F�C�v���Ŕw�ʂɂ���
Sub AM_�I�𒆃V�F�C�v���őO�ʂ�()
    If TypeName(Selection) = "Range" Then
        MsgBox "�V�F�C�v����I��ł�����s���ĂˁB"
        Exit Sub
    End If
    
    For Each shp In Selection.ShapeRange
        shp.ZOrder msoBringToFront
    Next
End Sub
'#�V�F�C�v�̏ꏊ�ɒl���Ȃ��Ȃ�悤�ɋ�s��}������
Sub AN_�V�F�C�v�\�t���u�����N�s�}��()
    '#�N���b�v�{�[�h�Ƀf�[�^�����鎞�̂�
    If Application.ClipboardFormats(1) Then
        '�\�t�BCtrl + V�ɂ�����A�N�V�����i���̎��_��Selection�̓V�F�C�v�ɂȂ�j
        ActiveSheet.Paste
        
        '�ړ��ʒu���擾���邽�߂̃_�~�[�V�F�C�v
        Dim dummyShape As Shape
        
        '�������̃Z�����擾���邽�߂̃_�~�[�V�F�C�v
        Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, Selection.top + Selection.Height, 1, 1)
    
        '�u�Z���ɍ��킹�Ĉړ���T�C�Y�ύX�����Ȃ��v�ɐݒ�
        '������Ȃ��ƍs�̑}���ɍ��킹�ăV�F�C�v���ꏏ�ɐL�тĂ��܂�����
        Selection.Placement = xlFreeFloating
        
        '�\�t�V�F�C�v�̉��ɂ���Z�������[�v
        For i = Selection.TopLeftCell.Row To dummyShape.TopLeftCell.Row + 1
            '�񃋁[�v�i��j=Selection.TopLeftCell.Column�Ȃ�V�F�C�v�\�t�ʒu����J�n�j
            For j = 1 To 15
                '�Ώۍs�̂ǂ����ɒl������΍s��}������
                If Cells(i, j) <> "" Then
                    Rows(i).Insert
                    Exit For
                End If
            Next
        Next
    
        '�p�ς݂�����폜����
        dummyShape.Delete
    End If
End Sub
'#�擪�ɖڎ��̃V�[�g���쐬����
Sub AO_�ڎ��V�[�g���쐬����()
    Dim ws As Worksheet
    
    '�֐��͕ʓr�Q��
    If Not isExistCheckToSheet(ActiveWorkbook, "�ڎ�") Then
        Worksheets.Add before:=Sheets(1)
        Set ws = Sheets(1)
        ws.Name = "�ڎ�"
        ws.Cells(1, 1) = "No."
        ws.Cells(1, 2) = "�V�[�g��"
        ws.Cells(1, 3) = "�V�[�g�̐���"
        ws.Cells(1, 4) = "�V�F�C�v�̐�"
        ws.Cells(1, 5) = "�g�p�͈�"
        ws.Cells(1, 6) = "���l"
        ws.Cells(1, 7) = "�쐬��"
        ws.Cells(1, 8) = "�쐬��"
        ws.Cells(1, 9) = "�X�V��"
        ws.Cells(1, 10) = "�X�V��"
        '�t�H���g�F
        Range("A1:J1").Font.Color = RGB(20, 10, 10)
        '�w�i�F
        Range("A1:J1").Interior.Color = RGB(255, 242, 204)
        '����
        Range("A1:J1").Font.Bold = True
        Cells(2, 1).Select
        '�E�B���h�E�g�̌Œ�
        ActiveWindow.FreezePanes = True
        '�ڐ�����\��
        ActiveWindow.DisplayGridlines = False
    Else
        Set ws = Sheets(1)
    End If
    
    Dim loopWs As Worksheet
    
    For i = 2 To Worksheets.count
        Set loopWs = Worksheets(i)
        ws.Cells(i, 1) = i - 1
        ws.Cells(i, 2) = loopWs.Name
        ws.Cells(i, 4) = loopWs.Shapes.count
        ws.Cells(i, 5) = loopWs.UsedRange.Address
        '�V�[�g�ł͂Ȃ��u�b�N�P�ʂ̏��Ȃ��߃R�����g�A�E�g
'        ws.Cells(i, 7) = ActiveWorkbook.BuiltinDocumentProperties(3)
'        ws.Cells(i, 8) = ActiveWorkbook.BuiltinDocumentProperties(11)
'        ws.Cells(i, 8).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
'        ws.Cells(i, 9) = ActiveWorkbook.BuiltinDocumentProperties(7)
'        ws.Cells(i, 10) = ActiveWorkbook.BuiltinDocumentProperties(12)
'        ws.Cells(i, 10).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    Next
    ws.Columns("A:H").AutoFit
    '���K�v������Έȉ����R�����g�C��
    ws.Cells(i + 1, 1) = "�K�v�ɉ����ĉ��L�̊֐���ǉ�����B�ڎ��V�[�g�ւ̃V���[�g�J�b�g�֐�"
    ws.Cells(i + 2, 1) = "Sub �ڎ��V�[�g��I��"
    ws.Cells(i + 3, 1) = "    Sheets(1).Select"
    ws.Cells(i + 4, 1) = "End Sub"
End Sub
'�w�肳�ꂽ�u�b�N�Ɏw�肳�ꂽ�V�[�g�����݂��邩�ǂ�����Ԃ��B���݂���:TRUE�A���݂��Ȃ�:FALSE
Function isExistCheckToSheet(wb As Workbook, checkSheet As Variant)
    isExistCheckToSheet = False
    For Each ws In wb.Worksheets
        If Not IsNumeric(checkSheet) Then
            If ws.Name = checkSheet Then
                isExistCheckToSheet = True
            End If
        End If
    Next
    
    If IsNumeric(checkSheet) Then
        '�w��l�����l�̏ꍇ�A�S�V�[�g����菬�����������ǂ��������݃`�F�b�N�ƂȂ�
        isExistCheckToSheet = checkSheet <= wb.Worksheets.count
    End If
End Function
'#�A�N�e�B�u�V�[�g�̓��e�ɏ]���V�[�g�𐶐����A�����N��t�^����B�\�[�g���s��
Sub AP_�V�[�g�����ƃ����N�t�^()

    '[�ڎ�]�V�[�g��ł̎��s��z�肵�Ă���B
    Dim topSheet As Worksheet
    Set topSheet = ActiveSheet
    
    '2��ڂɒl�̂���Ō�̍s���擾����
    Dim lastRowToBottom As Integer: lastRowToBottom = topSheet.Cells(1, 2).End(xlDown).Row
    
    Dim sheetName As String
    Dim linkRange As Range

    For i = 2 To lastRowToBottom
        sheetName = topSheet.Cells(i, 2).Value
        Set linkRange = topSheet.Cells(i, 2)
        
        If Not existsSheet(sheetName) Then
            '�V�[�g�����݂��Ă��Ȃ��ꍇ
            With Worksheets.Add(after:=ActiveSheet)
                '�V�[�g�𐶐����A�����N��t�^����
                .Name = sheetName
                topSheet.Hyperlinks.Add Anchor:=linkRange, Address:="", SubAddress:="'" & .Name & "'!A1"
                .Select
                '�ڐ�����\���A�V�[�g�k�ڒ���
                ActiveWindow.DisplayGridlines = False
                ActiveWindow.Zoom = 75
            End With
        Else
            '���ɃV�[�g������ꍇ
            topSheet.Hyperlinks.Add Anchor:=linkRange, Address:="", SubAddress:="'" & sheetName & "'!A1"
            Sheets(sheetName).Select
                '�ڐ�����\���A�V�[�g�k�ڒ���
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = 75
            Sheets(sheetName).Cells(1, 1).Select
            '�V�[�g���̕��ѕς�
            If existsSheet(topSheet.Cells(i - 1, 2)) Then
                '�����A�V�[�g���̎w��Ɂu.value�v���K�v
                Sheets(sheetName).move after:=Sheets(topSheet.Cells(i - 1, 2).Value)
            End If
        End If
    Next
    
    topSheet.Select
End Sub
'�V�[�g�����݂��邩�ǂ���
Function existsSheet(ByVal sheetName As String)
    Dim ws As Variant
    For Each ws In Sheets
        If LCase(ws.Name) = LCase(sheetName) Then
            existsSheet = True
            Exit Function
        End If
    Next

    '���݂��Ȃ�
    existsSheet = False
End Function
'�V�F�C�v�̂����I�𒆃Z����left�v���p�e�B�Ɉ�v���Ȃ����̂�I���Z���ʒu������ׂ�
Sub AQ_�V�F�C�v�ǉ�����()
Attribute AQ_�V�F�C�v�ǉ�����.VB_ProcData.VB_Invoke_Func = " \n14"
    '���摜�Ԃ̊Ԋu
    Const MARGIN_BOTTOM = 70
    
    
    '�\�t���W���i�[����itop�͓s�x���������Aleft�͏����l���g���܂킷�j
    Dim top As Integer: top = Selection.top + 5
    
    '�L���v�V�������L�ڂ���p�̃Z��
    Dim captionRange As Range
    Dim moveShape As Shape '�L���v�V�����^�C�g��
    Dim captionText As String: captionText = "��"
    For Each moveShape In ActiveSheet.Shapes
        '���ɊY�����Ȃ����̂͑ΏۊO�F�摜�A�O���[�v
        '�������͑I�𒆃Z��left�v���p�e�B�ƑΏۃV�F�C�vleft�v���p�e�B����v���Ȃ�����
        If (moveShape.Type <> msoPicture _
            And moveShape.Type <> msoGroup _
            And (moveShape.Type = msoAutoShape And Not moveShape.Fill.Visible)) _
            Or moveShape.left = Selection.left Then
            GoTo CONTINUE:
        End If
        
        '�V�F�C�v���ړ������āA
        Set captionRange = move(moveShape, top)
        
        '���L���v�V�������͂̐ݒ�i�s�v�Ȃ�R�����g�A�E�g���āj
        Call setCaption(captionRange, captionText)
        
        '���Ώۂɂ����V�F�C�v�̏㕔���W + ���Ώۂɂ����V�F�C�v�̍��� + �摜�Ԃ̊Ԋu + �L���v�V�����Z���s�̍��� = ���̃V�F�C�v�̈ړ���㕔���W
        top = top + moveShape.Height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).Height
CONTINUE:
    Next
End Sub
'�V�F�C�v����2�J�����ҁi���1��ځA������2��ځj
Sub AR_�V�F�C�v��2�J�����ŕ��ׂ�()
    '���摜�Ԃ̊Ԋu
    Const MARGIN_BOTTOM = 70
    
    '�\�t���W���i�[����itop�͓s�x���������Aleft�͏����l���g���܂킷�j
    Dim top As Integer: top = Selection.top + 5
    Dim left As Integer: left = Selection.left
    
    '�L���v�V�������L�ڂ���p�̃Z��
    Dim captionRange As Range
    Dim moveShape As Shape
    
    '�G���[�`�F�b�N
    If Selection.Row - REMARK_LINE - 1 < 1 Then
        MsgBox "�L���v�V�����p�̍s������܂���B����" & REMARK_LINE - Selection.Row + 2 & "�s���̈ʒu�Ŏ��s���Ă��������B"
        Exit Sub
    End If
    
    '�L���v�V�����^�C�g��
    Dim captionText As String
    '���_�C�A���O���g���ꍇ�͈ȉ��̃R�����g�A�E�g�����g�p����
    If CAPTION_TEXT_TOP_FLAG Then
        captionText = "���ύX�O���ύX��" 'InputBox("�L���v�V�����̏����l�����āB", "�L���v�V�����I�v�V����", "�������ɉ摜�̐���������")
    Else
        captionText = "���ύX�O���ύX��" 'InputBox("�L���v�V�����̏����l�����āB", "�L���v�V�����I�v�V����", "�������ɉ摜�̐���������")
    End If
    
    If StrPtr(answer) = 0 Then
        '�L�����Z����
        Exit Sub
    End If
    Dim count As Integer: count = 1
    For Each moveShape In ActiveSheet.Shapes
        '���ɊY�����Ȃ����̂͑ΏۊO�F�摜�A�O���[�v�A�h��Ԃ��̂Ȃ��I�[�g�V�F�C�v
        '�����ύX���Q�l�Fhttps://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
        If moveShape.Type <> msoPicture _
            And moveShape.Type <> msoGroup _
            And (moveShape.Type = msoAutoShape) Then  'And Not moveShape.Fill.Visible
            GoTo CONTINUE:
        End If
        
        '�ړ��ʒu���擾���邽�߂̃_�~�[�V�F�C�v
        Dim dummyShape As Shape
    
        '������̃Z�����擾���邽�߂̃_�~�[�V�F�C�v
        Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
        
        '�V�F�C�v���ړ�����
        moveShape.top = dummyShape.TopLeftCell.Offset(0, 0).top
        If count Mod 2 = 1 Then
            '���E�ɕ��ԃV�F�C�v�̍���
            moveShape.left = Selection.left
        ElseIf count Mod 2 = 0 Then
            '���E�ɕ��ԃV�F�C�v�̉E��
            moveShape.left = left
        End If
        
        moveShape.Name = "image-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
        
        '�L���v�V�������͗p�Z�����擾����i-1�̓^�C�g�����j
        Set captionRange = dummyShape.TopLeftCell.Offset(-1 - REMARK_LINE, 0)
        
        '�p�ς݂�����폜����
        dummyShape.Delete
        
        '���L���v�V�������͂̐ݒ�i�s�v�Ȃ�R�����g�A�E�g���āj
        Call setCaption(captionRange, captionText)
        
        If count Mod 2 = 1 Then
            'top�ɗ^���鐔�l��ς��Ȃ��B���̃V�F�C�v�ɗ^����left�v���p�e�B�l��ݒ肷��
            left = Selection.left + moveShape.WIDTH - 20
            '���ω����������B�s�v�Ȃ�R�����g�A�E�g���āB
            'Set onShape = ActiveSheet.Shapes.AddShape(msoShapeRightArrow, left - 10, top + moveShape.Height / 2, 40, 50)
        ElseIf count Mod 2 = 0 Then
            '���Ώۂɂ����V�F�C�v�̏㕔���W + ���Ώۂɂ����V�F�C�v�̍��� + �摜�Ԃ̊Ԋu + �L���v�V�����Z���s�̍��� = ���̃V�F�C�v�̈ړ���㕔���W
            top = top + moveShape.Height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).Height
        End If
        count = count + 1
CONTINUE:
    Next
    
    'END����
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
    Call setCaption(dummyShape.TopLeftCell, "- END -")
    dummyShape.Delete
    
End Sub
'�}�`�̖���t�^����
Sub AS_�V�F�C�v�Ԃɐ}�`����u��()
Attribute AS_�V�F�C�v�Ԃɐ}�`����u��.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim startShape As Shape
    Dim endShape As Shape
    Dim connectShape As Shape
    
    Dim x1 As Double
    Dim x2 As Double
    Dim y1 As Double
    Dim y2 As Double
    Dim degree As Double
    Dim adjustDegree As Integer
    Dim onShape As Shape
    Dim setLeft As Double
    Dim setTop As Double
    
    If TypeName(Selection) = "Range" Then
        MsgBox "�V�F�C�v���I������Ă��܂���B2�ȏ�I�����Ă��������B"
        Exit Sub
    End If
    For Each shp In Selection.ShapeRange
        If shp.Type = msoGroup Or shp.Connector Then
            MsgBox "�I���V�F�C�v�ɃO���[�v���R�l�N�^���܂܂�Ă��܂��B�������Ă��������B"
            Exit Sub
        End If
    Next
    
    For i = 1 To Selection.ShapeRange.count - 1
        '�I�𒆃V�F�C�v�̕ێ��i���j
        Set startShape = Selection.ShapeRange.Item(i)
        '�I�𒆃V�F�C�v�̕ێ��i��j
        Set endShape = Selection.ShapeRange.Item(i + 1)
        '�e�|�C���g���擾
        x1 = startShape.left + startShape.WIDTH - (startShape.WIDTH / 2)
        x2 = endShape.left + endShape.WIDTH - (endShape.WIDTH / 2)
        y1 = startShape.top + startShape.Height - (startShape.Height / 2)
        y2 = endShape.top + endShape.Height - (endShape.Height / 2)
        
        '���ʒu�v���p�e�B�̒���
        If startShape.left < endShape.left Then
            adjustDegree = 180
            setLeft = startShape.left + startShape.WIDTH + ((endShape.left - (startShape.left + startShape.WIDTH)) / 2) - 25
        Else
            setLeft = endShape.left + endShape.WIDTH + ((startShape.left - (endShape.left + endShape.WIDTH)) / 2) - 25
            
        End If
        
        '��ʒu�v���p�e�B�̒���
        If startShape.top < endShape.top Then
            setTop = startShape.top + (startShape.Height / 2) + (((endShape.top + (endShape.Height / 2)) - (startShape.top + (startShape.Height / 2))) / 2) - 25
        Else
            setTop = endShape.top + (endShape.Height / 2) + (((startShape.top + (startShape.Height / 2)) - (endShape.top + (endShape.Height / 2))) / 2) - 25
        End If
        '���̌��������B�Ō�Ɋ����Ă�͉̂~����
        If x2 - x1 <> 0 Then
            degree = Atn((y2 - y1) / (x2 - x1)) * 180 / 3.14
        Else
            degree = -90
        End If
        
        Set onShape = ActiveSheet.Shapes.AddShape(msoShapeLeftArrow, setLeft, setTop, 50, 50)
        onShape.Name = "allow-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
        onShape.Rotation = degree + adjustDegree
        adjustDegree = 0
    Next
End Sub
'�Z������Z���֘g�V�F�C�v���q������t�^����
Sub AT_�Z������Z���ɐL�т�R�l�N�^_�g����()
    Dim onShape As Shape
    
    For Each rcell In Selection
        Set onShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, _
                                                    rcell.MergeArea.left, _
                                                    rcell.MergeArea.top, _
                                                    rcell.MergeArea.WIDTH, _
                                                    rcell.MergeArea.Height)
        onShape.Name = "shape-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
        '���h��Ԃ��imsoTrue:����AmsoFalse:�Ȃ��j
        onShape.Fill.Visible = msoFalse
        '�����̑����B���D�݂łǂ���
        onShape.Line.Weight = 1
        '���F�w��
        onShape.Line.ForeColor.RGB = RGB(0, 0, 0)
        '���h��Ԃ��F�w��
        'onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
        '�R�l�N�^�Ōq�����߁A�I����Ԃɂ���
        onShape.Select Replace:=False
    Next
    '�V�F�C�v�Ԃɖ���t�^���Ă���
    Call AF_�V�F�C�v��I�����ɃR�l�N�^�Ōq��
End Sub
'�F��X�^�C����ύX�������Ƃ��A������Ɏ��{���A���Ƃ̓R�s�[����Ƃ�������
Sub AU_�ŏ��̃V�F�C�v���X�L�����R�s�[()
    '�����V�F�C�v�I�����A2�ڈȍ~��1�ڂ̃X�^�C����K�p����B���[�v����
    Dim baseShp As Shape
    Dim shp As Shape
    Set baseShp = Selection.ShapeRange.Item(1)
    For i = 2 To Selection.ShapeRange.count
        '�I�𒆃V�F�C�v�̕ێ��i�ڑ����j
        Set shp = Selection.ShapeRange.Item(i)
        shp.Line.ForeColor.RGB = baseShp.Line.ForeColor.RGB
        'shp.Line.Weight = baseShp.Line.Weight
        'shp.ForeColor.RGB = baseShp.ForeColor.RGB
        '�������[�h�A�[�g�t�H�[�}�b�g���w��ł��Ȃ��V�F�C�v��I�ԂƂ��̓R�����g�A�E�g����
        'shp.TextFrame2.WordArtformat = baseShp.TextFrame2.WordArtformat
        shp.Fill.Transparency = baseShp.Fill.Transparency
        '���e�L�X�g�܂ŕς������Ȃ��Ƃ��̓R�����g�A�E�g
        'shp.TextFrame.Characters.Text = baseShp.TextFrame.Characters.Text
        shp.Fill.ForeColor.RGB = baseShp.Fill.ForeColor.RGB
        shp.TextFrame2.TextRange.Font.Size = baseShp.TextFrame2.TextRange.Font.Size
        shp.TextFrame2.WordWrap = baseShp.TextFrame2.WordWrap
        shp.TextFrame.Characters.Font.Color = baseShp.TextFrame.Characters.Font.Color
        shp.TextFrame.Characters.Font.Name = baseShp.TextFrame.Characters.Font.Name
        shp.TextFrame2.VerticalAnchor = baseShp.TextFrame2.VerticalAnchor
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = baseShp.TextFrame2.TextRange.ParagraphFormat.Alignment
        shp.Placement = baseShp.Placement
        shp.LockAspectRatio = baseShp.LockAspectRatio
        shp.TextFrame2.AutoSize = baseShp.TextFrame2.AutoSize
        shp.TextFrame2.MarginLeft = baseShp.TextFrame2.MarginLeft
        shp.TextFrame2.MarginRight = baseShp.TextFrame2.MarginRight
        shp.TextFrame2.MarginTop = baseShp.TextFrame2.MarginTop
        shp.TextFrame2.MarginBottom = baseShp.TextFrame2.MarginBottom
        shp.TextFrame2.WordWrap = baseShp.TextFrame2.WordWrap
        shp.TextFrame2.VerticalAnchor = baseShp.TextFrame2.VerticalAnchor
        shp.TextFrame2.HorizontalAnchor = baseShp.TextFrame2.HorizontalAnchor
        shp.TextFrame2.Orientation = baseShp.TextFrame2.Orientation
        '���Ȃ������s���ɃG���[
        'shp.Shadow.Type = baseShp.Shadow.Type
        shp.Shadow.Visible = baseShp.Shadow.Visible
        '���Ȃ������s���ɃG���[
        'shp.Shadow.Style = baseShp.Shadow.Style
        shp.Shadow.Blur = baseShp.Shadow.Blur
        shp.Shadow.OffsetX = baseShp.Shadow.OffsetX
        shp.Shadow.OffsetY = baseShp.Shadow.OffsetY
        shp.Shadow.RotateWithShape = baseShp.Shadow.RotateWithShape
        shp.Shadow.ForeColor.RGB = baseShp.Shadow.ForeColor.RGB
        '���Ȃ������s���ɃG���[
        'shp.Shadow.Transparency = baseShp.Shadow.Transparency
        shp.Shadow.Size = baseShp.Shadow.Size
    Next
End Sub
'�ȗ����������B���A���A����3�{�̂ɂ��ɂ������쐬���A�Ō�ɃO���[�v�����Ă���
Sub AV_�ȗ��ɂ��ɂ��o��()
    Dim selectRange As Range
    Set selectRange = Selection
    Dim top As Integer: top = Selection.top
    Dim left As Integer: left = Selection.left
    '���ɂ��ɂ��̒�����ݒ肷��B�������傫���ƌ��\���Ԃ�������
    Const WIDTH = 50
    Dim blackTopShape As Shape
    Dim whiteShape As Shape
    Dim blackBottomShape As Shape

    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, left, top)
        For i = 0 To WIDTH
            If i Mod 2 = 0 Then
                top = top + 5
            Else
                top = top - 5
            End If
            left = left + 7
            .AddNodes msoSegmentCurve, msoEditingAuto, left, top
        Next
        Set blackTopShape = .ConvertToShape
    End With
    blackTopShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    blackTopShape.Line.Weight = 3
    blackTopShape.Name = "omit-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()

    top = top + 3
    left = Selection.left
    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, left, top)
    
        For i = 0 To WIDTH
            If i Mod 2 = 0 Then
                top = top + 5
            Else
                top = top - 5
            End If
            left = left + 7
            .AddNodes msoSegmentCurve, msoEditingAuto, left, top
        Next
        Set blackBottomShape = .ConvertToShape
    End With
    blackBottomShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    blackBottomShape.Line.Weight = 3
    blackBottomShape.Name = "omit-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
    
    
    top = top - 9
    left = Selection.left
    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, left, top)
    
        For i = 0 To WIDTH
            If i Mod 2 = 0 Then
                top = top + 5.2
            Else
                top = top - 5.2
            End If
            left = left + 7
            .AddNodes msoSegmentCurve, msoEditingAuto, left, top
        Next
        Set whiteShape = .ConvertToShape
    End With
    whiteShape.Line.ForeColor.RGB = RGB(255, 255, 255)
    whiteShape.Line.Weight = 5.8
    whiteShape.Name = "omit-" & Format(Now(), "yyyymmdd-hhmmss.") & getMSec()
    
    blackTopShape.Select Replace:=False
    whiteShape.Select Replace:=False
    blackBottomShape.Select Replace:=False
    Selection.Group
    selectRange.Select
End Sub
'�F���m�F���܂��˂�
Sub AW_�F�m�F()
Attribute AW_�F�m�F.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim fontColorCode As Long
    Dim backColorCode As Long
    Dim lineColorCode As Long

    If TypeName(Selection) = "Range" Then
        Debug.Print "���Z���̏��i" & Now() & "�j������������-��"
        '�Z���̔w�i�F
        backColorCode = Selection.Interior.Color
        '�Z���̕����F
        fontColorCode = Selection.Font.Color
        '�Z���̘g���F
        lineColorCode = Selection.Borders.Color
        
    ElseIf Selection.ShapeRange.Connector Then
        Debug.Print "�����̏��i" & Now() & "�j������������-��"
        '���̘g���F
        lineColorCode = Selection.ShapeRange.Item(1).Line.ForeColor.RGB
        
    ElseIf TypeName(Selection) = "Rectangle" Then
        Debug.Print "���V�F�C�v�̏��i" & Now() & "�j��������-��"
        '�V�F�C�v�̓h��Ԃ��F
        backColorCode = Selection.ShapeRange.Item(1).Fill.ForeColor.RGB
        '�V�F�C�v�̕����F
        fontColorCode = Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
        '�V�F�C�v�̘g���F
        lineColorCode = Selection.ShapeRange.Item(1).Line.ForeColor.RGB
    Else
        Debug.Print "�������ɓ������ꍇ�z��O"
    End If

    Debug.Print "�� �w�i�F/�h��Ԃ��F�̐F�R�[�h�F" & backColorCode
    Dim Red As Integer: Red = backColorCode Mod 256
    Dim Green As Integer: Green = Int(backColorCode / 256) Mod 256
    Dim Blue As Integer: Blue = Int(backColorCode / 256 / 256)
    Debug.Print "�� �w�i�F/�h��Ԃ��F��RGB�l�FRGB(" & Red & "," & Green; "," & Blue & ")"
    Debug.Print "��������������������������������������������������"
    Debug.Print "�� �����F�̐F�R�[�h�F" & fontColorCode
    Red = fontColorCode Mod 256
    Green = Int(fontColorCode / 256) Mod 256
    Blue = Int(fontColorCode / 256 / 256)
    Debug.Print "�� �����F��RGB�l�FRGB(" & Red & "," & Green; "," & Blue & ")"
    Debug.Print "��������������������������������������������������"
    Debug.Print "�� �Z���g�F/���F�̐F�R�[�h�F" & lineColorCode
    Red = lineColorCode Mod 256
    Green = Int(lineColorCode / 256) Mod 256
    Blue = Int(lineColorCode / 256 / 256)
    Debug.Print "�� �Z���g�F/���F��RGB�l�FRGB(" & Red & "," & Green; "," & Blue & ")"
    Debug.Print "��������������������������������������������������"

End Sub
Sub AX_X���W���킹()
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).left + Selection.ShapeRange.Item(1).WIDTH / 2
    For Each sp In Selection.ShapeRange
        sp.left = baseXCenterPoint - sp.WIDTH / 2
    Next
End Sub
Sub AY_Y���W���킹()
    Dim baseXCenterPoint As Integer: baseXCenterPoint = Selection.ShapeRange.Item(1).top + Selection.ShapeRange.Item(1).Height / 2
    For Each sp In Selection.ShapeRange
        sp.top = baseXCenterPoint - sp.Height / 2
    Next
End Sub
'�Z������Z���֖�������B1��2,3��4�̂悤�ɂ���
Sub AZ_�Z������Z���ɐL�т�R�l�N�^_�g�Ȃ�()
    Dim count As Integer: count = 1
    '�n�_�Z��
    Dim startRange As Range
    '�I�_�Z��
    Dim endRange As Range
    '���V�F�C�v
    Dim connectShape As Shape
    For Each cell In Selection
        If count Mod 2 = 1 Then
            '��̎��͎n�_�Z����ϐ���
            Set startRange = cell
        ElseIf count Mod 2 = 0 Then
            '�����̎��͏I�_�Z���̐ݒ�Ɩ��̑}��
            Set endRange = cell
            Set connectShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorElbow, _
                BeginX:=startRange.left + (startRange.WIDTH), BeginY:=startRange.top + (startRange.Height / 2), _
                EndX:=endRange.left + (endRange.WIDTH / 2), EndY:=endRange.top + (endRange.Height / 2))
            Call makeConnectAllow(connectShape)
        End If
        count = count + 1
    Next
End Sub
'�w��F���ŐF�̐؂�ւ����s���B�K�v�ɉ����ăR�����g�A�E�g�̈ʒu�A�F�ݒ�v���p�e�B�𒲐�����
Sub BA_�F�؂�ւ�()
Attribute BA_�F�؂�ւ�.VB_ProcData.VB_Invoke_Func = " \n14"
    '�V�F�C�v�̓h��Ԃ��F
    'Dim colorCode As String: colorCode = Selection.ShapeRange.Item(1).Fill.ForeColor.RGB
    '�V�F�C�v�̕����F
    'Dim colorCode As String: colorCode = Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    '�V�F�C�v�̘g���F
    Dim colorCode As String: colorCode = Selection.ShapeRange.Item(1).Line.ForeColor.RGB

    Dim colors() As Variant
    '���F�R�[�h���X�g�i�uAW_�F�m�F�v�ŃR�[�h���擾���ݒ肷��j
    colors = Array("255", "49407", "65535", "5296274", "5287936", "15773696", "12611584")
    
    Dim hitIndex As Integer
    hitIndex = isExistArrayReturnIndex(colors, colorCode)
    
    If hitIndex <> -1 Then
        If UBound(colors) = hitIndex Then
            '�F�z��̍ŏI�C���f�b�N�X�������ꍇ�A0�Ԗڂɖ߂�
            Selection.ShapeRange.Item(1).Line.ForeColor.RGB = colors(0)
        Else
            '�F�z��ɑ����Ă���A���̃C���f�b�N�X�F�ɐݒ肷��
            Selection.ShapeRange.Item(1).Line.ForeColor.RGB = colors(hitIndex + 1)
        End If
    Else
        '�F�z��̂ǂ�ɂ�������Ȃ������ꍇ
        Selection.ShapeRange.Item(1).Line.ForeColor.RGB = colors(0)
    End If
    
End Sub
