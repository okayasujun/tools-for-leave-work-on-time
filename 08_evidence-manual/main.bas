Attribute VB_Name = "main"
'###############################
'�@�\���F�G�r�f���X�E�}�j���A���쐬�x���c�[�� v2.0
'Author�Fokayasu jun
'�쐬���F2022/10/19
'�X�V���F2023/02/25
'COMMENT�F
'###############################
'�|�C���^API�B�}�E�X�J�[�\���ʒu����Z���ʒu���擾���邽�߂Ɏg�p����
Private Type POINTAPI
    x As Long
    y As Long
End Type
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'���摜�^�C�g���ɕ�L������ꍇ�̍s���i�^�C�g���Ɖ摜�̊Ԃ̋�s���j
Const REMARK_LINE = 0
'#�I���Z���͈͂̑傫���̐Ԙg���}�E�X�ʒu�ɏo��������
Sub A_�Ԙg���o��������()
Attribute A_�Ԙg���o��������.VB_ProcData.VB_Invoke_Func = "q\n14"

    Dim onShape As shape
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
                                            Selection.width, _
                                            Selection.height)
    '���h��Ԃ��imsoTrue:����AmsoFalse:�Ȃ��j
    onShape.Fill.Visible = msoFalse
    '�����̑����B���D�݂łǂ���
    onShape.Line.Weight = 4
    '���F�w��
    onShape.Line.ForeColor.RGB = RGB(255, 0, 0)
    '���h��Ԃ��F�w��
    'onShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
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

    '�}�E�X�J�[�\���̈ʒu����Z�����擾�i�J�[�\���̏�Ԏ���ł͎��s����j
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
Sub B_�e��t����()
Attribute B_�e��t����.VB_ProcData.VB_Invoke_Func = "w\n14"

    Dim shp As shape

    If TypeName(Selection) = "Range" Then
        '�V�F�C�v���I����ԁB�S�V�F�C�v��Ώۂɂ���
        For Each shp In ActiveSheet.Shapes
            '�����ύX���Q�l�Fhttps://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
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
'�w�肳�ꂽ�V�F�C�v�ɉe��t�^���܂��B
Function castShadow(shp As shape)
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
Sub C_���ʂ����Z�b�g����()
Attribute C_���ʂ����Z�b�g����.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim shp As shape
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
Function shapeReset(shp As shape)
    If shp.Type = msoPicture Or (shp.Type = msoAutoShape And shp.Fill.Visible) Then
        With Application.CommandBars
            '�u�}�̃��Z�b�g�v���\�ȂƂ��̂�
            If .GetEnabledMso("PictureReset") Then
                .ExecuteMso "PictureReset"
            End If
            '�I�[�g�V�F�C�v�̏ꍇ�͈ȉ��ŊO���i���P�[�X�ɉ����ď�����������ǉ����邱�Ɓj
            shp.Shadow.Visible = msoFalse
        End With
    End If
End Function
'#�V�F�C�v�𐮗񂳂���i�\�t���j
Sub D_�\�t���ɐ��񂳂���()
Attribute D_�\�t���ɐ��񂳂���.VB_ProcData.VB_Invoke_Func = "e\n14"
    '���摜�Ԃ̊Ԋu
    Const MARGIN_BOTTOM = 70
    
    
    '�\�t���W���i�[����itop�͓s�x���������Aleft�͏����l���g���܂킷�j
    Dim top As Integer: top = Selection.top + 5
    
    '�L���v�V�������L�ڂ���p�̃Z��
    Dim captionRange As Range
    Dim moveShape As shape
    
    '�G���[�`�F�b�N
    If Selection.Row - REMARK_LINE - 1 < 1 Then
        MsgBox "�L���v�V�����p�̍s������܂���B����" & REMARK_LINE - Selection.Row + 2 & "�s���̈ʒu�Ŏ��s���Ă��������B"
        Exit Sub
    End If
    
    '�L���v�V�����^�C�g��
    Dim captionText As String
    '���_�C�A���O���g���ꍇ�͈ȉ��̃R�����g�A�E�g�����g�p����
    captionText = "��" 'InputBox("�L���v�V�����̏����l�����āB", "�L���v�V�����I�v�V����", "�������ɉ摜�̐���������")
    
    If StrPtr(answer) = 0 Then
        '�L�����Z����
        Exit Sub
    End If
    
    For Each moveShape In ActiveSheet.Shapes
        '���ɊY�����Ȃ����̂͑ΏۊO�F�摜�A�O���[�v�A�h��Ԃ��̂Ȃ��I�[�g�V�F�C�v
        '�����ύX���Q�l�Fhttps://learn.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
        If moveShape.Type <> msoPicture _
            And moveShape.Type <> msoGroup _
            And (moveShape.Type = msoAutoShape And Not moveShape.Fill.Visible) Then
            GoTo CONTINUE:
        End If
        
        '�V�F�C�v���ړ������āA
        Set captionRange = move(moveShape, top)
        
        '���L���v�V�������͂̐ݒ�i�s�v�Ȃ�R�����g�A�E�g���āj
        Call setCaption(captionRange, captionText)
        
        '���Ώۂɂ����V�F�C�v�̏㕔���W + ���Ώۂɂ����V�F�C�v�̍��� + �摜�Ԃ̊Ԋu + �L���v�V�����Z���s�̍��� = ���̃V�F�C�v�̈ړ���㕔���W
        top = top + moveShape.height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).height
CONTINUE:
    Next
    
    'END����
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
    Call setCaption(dummyShape.TopLeftCell, "END")
    dummyShape.Delete
    
End Sub
'
Function move(moveShape As shape, top As Integer)
    '�ړ��ʒu���擾���邽�߂̃_�~�[�V�F�C�v
    Dim dummyShape As shape
    Dim left As Integer: left = Selection.left
    
    '������̃Z�����擾���邽�߂̃_�~�[�V�F�C�v
    Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, top, 1, 1)
        
    '�V�F�C�v���ړ�����
    moveShape.top = dummyShape.TopLeftCell.Offset(0, 0).top
    moveShape.left = Selection.left
        
    '�L���v�V�������͗p�Z�����擾����i-1�̓^�C�g�����j
    Set move = dummyShape.TopLeftCell.Offset(-1 - REMARK_LINE, 0)
        
    '�p�ς݂�����폜����
    dummyShape.Delete
End Function
'�L���v�V�����p�Z���̐ݒ�
Function setCaption(captionRange As Range, captionText As String)
    '�摜�Ԉړ���Ctrl+���ō����ɍs������
    captionText = IIf(captionText = "", " ", captionText)
    '���K�X�ς��Ă悵
    captionRange.Value = captionText
    captionRange.Font.Bold = True
    captionRange.Font.Color = RGB(0, 0, 0)
End Function
'�I�𒆂̃V�F�C�v��I�����ɃR�l�N�^�Ōq��
Sub F_�V�F�C�v��I�����ɃR�l�N�^�Ōq��()
Attribute F_�V�F�C�v��I�����ɃR�l�N�^�Ōq��.VB_ProcData.VB_Invoke_Func = "l\n14"
    Dim startShape As shape
    Dim endShape As shape
    Dim connectShape As shape
    
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
    
    For i = 1 To Selection.ShapeRange.Count - 1
        '�I�𒆃V�F�C�v�̕ێ��i�ڑ����j
        Set startShape = Selection.ShapeRange.Item(i)
        '�I�𒆃V�F�C�v�̕ێ��i�ڑ���j
        Set endShape = Selection.ShapeRange.Item(i + 1)

        '�ڑ��V�F�C�v�̒a��
        '��Type�����͉E�L���Q�ƁFhttps://learn.microsoft.com/ja-jp/office/vba/api/office.msoconnectortype
        Set connectShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorElbow, BeginX:=0, BeginY:=0, EndX:=0, EndY:=0)
        '���ڑ��̎n�_�ʒu�w��i�Ō�̈�����1:��ӁA2:���ӁA3:���ӁA4�E�Ӂj
        connectShape.ConnectorFormat.BeginConnect ActiveSheet.Shapes(startShape.Name), 3
        '���ڑ��̏I�_�ʒu�w��i�Ō�̈����͎n�_�ʒu�̎w��Ɠ��l�j
        connectShape.ConnectorFormat.EndConnect ActiveSheet.Shapes(endShape.Name), 2
        '���I�_�R�l�N�^���O�p�ɁB
        connectShape.Line.EndArrowheadStyle = msoArrowheadTriangle
        '�����̐F
        connectShape.Line.ForeColor.RGB = RGB(0, 0, 0)
        '�����̑���
        connectShape.Line.Weight = 1
        '���I�_�̒���
        connectShape.Line.EndArrowheadLength = msoArrowheadLong
        '���I�_�̑���
        connectShape.Line.EndArrowheadWidth = msoArrowheadWide
    Next
End Sub
'�R�l�N�^�̎�ނ𒼐��ƃG���{�[�Ő؂�ւ���
Sub G_�R�l�N�^��ސ؂�ւ�()
Attribute G_�R�l�N�^��ސ؂�ւ�.VB_ProcData.VB_Invoke_Func = "i\n14"

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
'�I�������R�l�N�^�̎n�_���̃V�F�C�v�Ƃ̐ڑ��ʒu��ύX����
Sub H_�R�l�N�^�n�_�ύX()
Attribute H_�R�l�N�^�n�_�ύX.VB_ProcData.VB_Invoke_Func = "o\n14"

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
Sub I_�R�l�N�^�I�_�ύX()
Attribute I_�R�l�N�^�I�_�ύX.VB_ProcData.VB_Invoke_Func = "p\n14"
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
Sub J_�I�𒆃V�F�C�v���O���[�v��()
    Selection.Group.Select
End Sub
'#�I�𒆂̃V�F�C�v���O���[�v����
Sub K_�I�𒆃V�F�C�v���O���[�v����()
    Selection.Ungroup
End Sub
'�I�𒆂̃V�F�C�v���Ŕw�ʂɂ���
Sub L_�I�𒆃V�F�C�v���Ŕw�ʂ�()
    If TypeName(Selection) = "Range" Then
        MsgBox "�V�F�C�v����I��ł�����s���ĂˁB"
        Exit Sub
    End If
    
    For Each shp In Selection.ShapeRange
        shp.ZOrder msoSendToBack
    Next
End Sub
'#�V�F�C�v�̏ꏊ�ɒl���Ȃ��Ȃ�悤�ɋ�s��}������
Sub N_�V�F�C�v�\�t����u�����N�s�}��()
    '#�N���b�v�{�[�h�Ƀf�[�^�����鎞�̂�
    If Application.ClipboardFormats(1) Then
        '�\�t�BCtrl + V�ɂ�����A�N�V�����i���̎��_��Selection�̓V�F�C�v�ɂȂ�͗l�j
        ActiveSheet.Paste
        '�ړ��ʒu���擾���邽�߂̃_�~�[�V�F�C�v
        Dim dummyShape As shape
        '�������̃Z�����擾���邽�߂̃_�~�[�V�F�C�v
        Set dummyShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.left, Selection.top + Selection.height, 1, 1)
    
        '�u�Z���ɍ��킹�Ĉړ���T�C�Y�ύX�����Ȃ��v�ɐݒ�
        '������Ȃ��ƍs�̑}���ɍ��킹�ăV�F�C�v���ꏏ�ɐL�тĂ��܂�����
        Selection.Placement = xlFreeFloating
        
        '�\�t�V�F�C�v�̉��ɂ���Z�������[�v
        For i = Selection.TopLeftCell.Row To dummyShape.TopLeftCell.Row
            '�񃋁[�v�i��j=Selection.TopLeftCell.Column�Ȃ�V�F�C�v�\�t�ʒu����J�n�j
            For j = 1 To 100
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
Sub O_�ڎ��V�[�g���쐬����()
    Dim ws As Worksheet
    
    '�֐��͕ʓr�Q��
    If Not isExistCheckToSheet(ThisWorkbook, "�ڎ�") Then
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
        '�t�H���g�F
        Range("A1:H1").Font.Color = RGB(20, 10, 10)
        '�w�i�F
        Range("A1:H1").Interior.Color = RGB(255, 242, 204)
        '����
        Range("A1:H1").Font.Bold = True
        Cells(2, 1).Select
        '�E�B���h�E�g�̌Œ�
        ActiveWindow.FreezePanes = True
        '�ڐ�����\��
        ActiveWindow.DisplayGridlines = False
    Else
        Set ws = Sheets(1)
    End If
    
    Dim loopWs As Worksheet
    
    For i = 2 To Worksheets.Count
        Set loopWs = Worksheets(i)
        ws.Cells(i, 1) = i - 1
        ws.Cells(i, 2) = loopWs.Name
        ws.Cells(i, 4) = loopWs.Shapes.Count
        ws.Cells(i, 5) = loopWs.UsedRange.Address
    Next
    ws.Columns("A:H").AutoFit
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
        isExistCheckToSheet = checkSheet <= wb.Worksheets.Count
    End If
End Function
'#�A�N�e�B�u�V�[�g�̓��e�ɏ]���V�[�g�𐶐����A�����N��t�^����B�\�[�g���s��
Sub P_�V�[�g�����ƃ����N�t�^()
    'TODO:�ڐ����폜�A�k�ړ���
    Dim topSheet As Worksheet
    Set topSheet = ActiveSheet
    
    Dim lastRowToBottom As Integer: lastRowToBottom = topSheet.Cells(1, 2).End(xlDown).Row
    
    Dim sheetName As String
    Dim linkRange As Range

    For i = 2 To lastRowToBottom
        sheetName = topSheet.Cells(i, 2).Value
        Set linkRange = topSheet.Cells(i, 2)
        
        If Not existsSheet(sheetName) Then
            '�V�[�g�����݂��Ă��Ȃ��ꍇ
            With Worksheets.Add(after:=ActiveSheet)
                .Name = sheetName
                topSheet.Hyperlinks.Add Anchor:=linkRange, Address:="", SubAddress:=.Name & "!A1"
                .Select
            End With
        Else
            '���ɂ���ꍇ
            topSheet.Hyperlinks.Add Anchor:=linkRange, Address:="", SubAddress:=sheetName & "!A1"
            Sheets(sheetName).Select
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
Sub Q_�V�F�C�v�ǉ�����()
Attribute Q_�V�F�C�v�ǉ�����.VB_ProcData.VB_Invoke_Func = " \n14"
    '���摜�Ԃ̊Ԋu
    Const MARGIN_BOTTOM = 70
    
    
    '�\�t���W���i�[����itop�͓s�x���������Aleft�͏����l���g���܂킷�j
    Dim top As Integer: top = Selection.top + 5
    
    '�L���v�V�������L�ڂ���p�̃Z��
    Dim captionRange As Range
    Dim moveShape As shape '�L���v�V�����^�C�g��
    Dim captionText As String: captionText = "��"
    For Each moveShape In ActiveSheet.Shapes
        '���ɊY�����Ȃ����̂͑ΏۊO�F�摜�A�O���[�v�A�h��Ԃ��̂Ȃ��I�[�g�V�F�C�v�A
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
        top = top + moveShape.height + MARGIN_BOTTOM + Range(captionRange, captionRange.Offset(REMARK_LINE, 0)).height
CONTINUE:
    Next
End Sub
Sub R_2��3�J�����̕���()
    '�܂����x��������
End Sub
Sub U_�Z������Z���ɐL�т�R�l�N�^()
    '�܂����x��������
End Sub
'��������I�N�I���e�B�B�ߕs���͌���C��
Sub V_�ŏ��̃V�F�C�v���X�L�����R�s�[()
    '�����V�F�C�v�I�����A2�ڈȍ~��1�ڂ̃X�^�C����K�p����B���[�v����
    Dim baseShp As shape
    Dim shp As shape
    Set baseShp = Selection.ShapeRange.Item(1)
    For i = 2 To Selection.ShapeRange.Count
        '�I�𒆃V�F�C�v�̕ێ��i�ڑ����j
        Set shp = Selection.ShapeRange.Item(i)
        shp.Line.ForeColor.RGB = baseShp.Line.ForeColor.RGB
        'shp.ForeColor.RGB = baseShp.ForeColor.RGB
        '�����[�h�A�[�g�t�H�[�}�b�g���w��ł��Ȃ��V�F�C�v��I�ԂƂ��̓R�����g�A�E�g����
        shp.TextFrame2.WordArtformat = baseShp.TextFrame2.WordArtformat
        shp.Fill.Transparency = baseShp.Fill.Transparency
        '�e�L�X�g�܂ŕς������Ȃ��Ƃ��̓R�����g�A�E�g
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
    Next
End Sub
