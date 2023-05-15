Attribute VB_Name = "main"
'�J�n�s
Dim fromRow As Integer
'�J�n��
Dim fromColumn As Integer
'�ŏI�s
Dim toRow As Integer
'�ŏI��
Dim toColumn As Integer
'�h��Ԃ��Ȃ��������F�R�[�h
Const NO_COLOR = 16777215
'���������������F�R�[�h
Const INIT_TEXT_COLOR = 0
'�Z���̏����ݒ�̕W��
Const BASE_FORMAT = "G/�W��"
'��ԍČ��\�[�X�𐶐�����
Sub createRevivalSource()
    '�g�p�͈̓A�h���X
    Dim usedRangeAddress As String
    '�g�p�͈͊J�n�A�h���X
    Dim fromAddress As String
    '�g�p�͈͍ŏI�A�h���X
    Dim toAddress As String
    '�͈̓t���O�i���[�v�̍ۂ�2�d�ł�邩�ǂ����B�Ă�����g���Ă˂��ȁj
    Dim doubleLoopFlag As Boolean: doubleLoopFlag = False
    
    '�g�p�͈͎擾�i�l���Ȃ������������ΏۂƂȂ�j
    usedRangeAddress = ActiveSheet.UsedRange.Address
    'Debug.Print address
    
    If usedRangeAddress Like "*:*" Then
        '�g�p�͈͏��̎擾
        fromAddress = Split(usedRangeAddress, ":")(0)
        toAddress = Split(usedRangeAddress, ":")(1)
        'Debug.Print Range(fromAddress).Row
        'Debug.Print Range(fromAddress).Column
        'Debug.Print Range(toAddress).Row
        'Debug.Print Range(toAddress).Column
        fromRow = Range(fromAddress).Row
        fromColumn = Range(fromAddress).Column
        toRow = Range(toAddress).Row
        toColumn = Range(toAddress).Column
        doubleLoopFlag = True

    Else
        'Debug.Print Range(usedRangeAddress).Row
        'Debug.Print Range(usedRangeAddress).Column
        fromRow = Range(usedRangeAddress).Row
        toRow = Range(usedRangeAddress).Column

    End If
    
    '��ԍĐ����܂��B
    Call writeFromExcelToText
End Sub
'��ԍĐ��\�[�X����
Function writeFromExcelToText()
    Dim filePath As String ': filePath = ActiveWorkbook.Path & "\setup.bas"
    Dim moduleName As String
    moduleName = InputBox("�t�@�C���������āB�g���q�͂���Ȃ��B", "���W���[���t�@�C����", "setupX")
    filePath = ActiveWorkbook.Path & "\" & moduleName & ".bas"
    
    If StrPtr(moduleName) = 0 Then
        '�L�����Z����
        Exit Function
    End If
    
    Const CHAR_SET = "SHIFT-JIS" 'UTF-8 / SHIFT-JIS
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim count As Integer
    Dim functionCount As Integer: functionCount = 1
    Dim temp As String
    

    With CreateObject("ADODB.Stream")
        .Charset = CHAR_SET
        'https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/lineseparatorsenum?view=sql-server-ver16
        .LineSeparator = 10
        .Open
        '�t�@�C�������o��
        .WriteText "Attribute VB_Name = """ & moduleName & """", 1
        .WriteText "Function revival0()", 1
        .WriteText "    ActiveWindow.DisplayGridlines = " & ActiveWindow.DisplayGridlines, 1
        .WriteText "    ActiveWindow.Zoom = " & ActiveWindow.Zoom, 1
        .WriteText "    ActiveSheet.Name = """ & ActiveSheet.Name & """", 1
        

        For i = fromRow To toRow
            For j = fromColumn To toColumn
                '�l
                If ws.Cells(i, j).Value <> "" Then
                    .WriteText "    Cells(" & i & ", " & j & ") = """ & Replace(ws.Cells(i, j), vbLf, """& vbLf & """) & """", 1
                End If
                '����
                If ws.Cells(i, j).NumberFormatLocal <> BASE_FORMAT Then
                    .WriteText "    Cells(" & i & ", " & j & ").NumberFormatLocal = """ & ws.Cells(i, j).NumberFormatLocal & """", 1
                End If
                '�܂�Ԃ�
                If ws.Cells(i, j).WrapText Then
                    .WriteText "    Cells(" & i & ", " & j & ").WrapText  = " & ws.Cells(i, j).WrapText, 1
                End If
                '�t�H���g�T�C�Y
                If ws.Cells(i, j).Font.Size <> 11 Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Size = " & ws.Cells(i, j).Font.Size, 1
                End If
                '�t�H���g��
                If ws.Cells(i, j).Font.Name <> "" Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Name = """ & ws.Cells(i, j).Font.Name & """", 1
                End If
                '�w�i�F
                If ws.Cells(i, j).Interior.color <> NO_COLOR Then
                    .WriteText "    Cells(" & i & ", " & j & ").Interior.Color = " & ws.Cells(i, j).Interior.color, 1
                End If
                '�����F
                If ws.Cells(i, j).Font.color <> INIT_TEXT_COLOR Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Color = " & ws.Cells(i, j).Font.color, 1
                End If
                '����
                If ws.Cells(i, j).Font.Bold Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Bold = " & ws.Cells(i, j).Font.Bold, 1
                End If
                '�C�^���b�N
                If ws.Cells(i, j).Font.Italic Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Italic = " & ws.Cells(i, j).Font.Italic, 1
                End If
                '�����
                If ws.Cells(i, j).Font.Strikethrough Then
                    .WriteText "    Cells(" & i & ", " & j & ").Font.Strikethrough = " & ws.Cells(i, j).Font.Strikethrough, 1
                End If
                '�����ʒu
                If ws.Cells(i, j).HorizontalAlignment <> xlGeneral Then
                    .WriteText "    Cells(" & i & ", " & j & ").HorizontalAlignment = " & ws.Cells(i, j).HorizontalAlignment, 1
                End If
                '�����ʒu
                If ws.Cells(i, j).VerticalAlignment <> xlCenter Then
                    .WriteText "    Cells(" & i & ", " & j & ").VerticalAlignment = " & ws.Cells(i, j).VerticalAlignment, 1
                End If
                '�C���f���g���x��
                If ws.Cells(i, j).IndentLevel > 0 Then
                    .WriteText "    Cells(" & i & ", " & j & ").IndentLevel = " & ws.Cells(i, j).IndentLevel, 1
                End If
                '�Z���̃}�[�W
                If ws.Cells(i, j).MergeCells Then
                    .WriteText "    Range(""" & ws.Cells(i, j).MergeArea.Item(1).Address(0, 0) & ":" & ws.Cells(i, j).MergeArea.Item(ws.Cells(i, j).MergeArea.count).Address(0, 0) & """).Merge", 1
                End If
                '�r���i��j
                If ws.Cells(i, j).Borders(xlEdgeTop).LineStyle <> xlLineStyleNone Then
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeTop).LineStyle = " & ws.Cells(i, j).Borders(xlEdgeTop).LineStyle, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeTop).color = " & ws.Cells(i, j).Borders(xlEdgeTop).color, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeTop).weight = " & ws.Cells(i, j).Borders(xlEdgeTop).weight, 1
                End If
                '�r���i���j
                If ws.Cells(i, j).Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone Then
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeBottom).LineStyle = " & ws.Cells(i, j).Borders(xlEdgeBottom).LineStyle, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeBottom).color = " & ws.Cells(i, j).Borders(xlEdgeBottom).color, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeBottom).weight = " & ws.Cells(i, j).Borders(xlEdgeBottom).weight, 1
                End If
                '�r���i���j
                If ws.Cells(i, j).Borders(xlEdgeLeft).LineStyle <> xlLineStyleNone Then
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeLeft).LineStyle = " & ws.Cells(i, j).Borders(xlEdgeLeft).LineStyle, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeLeft).color = " & ws.Cells(i, j).Borders(xlEdgeLeft).color, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeLeft).weight = " & ws.Cells(i, j).Borders(xlEdgeLeft).weight, 1
                End If
                '�r���i�E�j
                If ws.Cells(i, j).Borders(xlEdgeRight).LineStyle <> xlLineStyleNone Then
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeRight).LineStyle = " & ws.Cells(i, j).Borders(xlEdgeRight).LineStyle, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeRight).color = " & ws.Cells(i, j).Borders(xlEdgeRight).color, 1
                    .WriteText "    Cells(" & i & ", " & j & ").Borders(xlEdgeRight).weight = " & ws.Cells(i, j).Borders(xlEdgeRight).weight, 1
                End If
                '���͋K���i���X�g�j
                On Error Resume Next '��ݒ莞�̃G���[���
                '�ϐ��i�[�͕s�v�����ǁA�A�N�Z�X���邱�ƂŐݒ�ۂ����Ԃ肾���B����ɉ�����ErrorNumber��]������
                temp = ws.Cells(i, j).Validation.Type
                If Err.Number = 0 Then
                    .WriteText "    Cells(" & i & ", " & j & ").Validation.Delete", 1
                    .WriteText "    Cells(" & i & ", " & j & ").Validation.Add Type:=xlValidateList, _", 1
                    .WriteText "        Operator:=xlEqual, _", 1
                    .WriteText "        Formula1:=""" & ws.Cells(i, j).Validation.Formula1 & """", 1
                End If
                Err.Clear
                '�֐���؂�i���s���ɋN����u�v���V�[�W�����傫�����܂��B�v��������邽�߁j
                If count > 30 And count Mod 30 = 0 Then
                    .WriteText "end Function", 1
                    .WriteText "Function revival" & functionCount & "()", 1
                    functionCount = functionCount + 1
                End If
                count = count + 1
            Next
        Next
        For i = fromRow To toRow
            '�s���ݒ�
            .WriteText "    Rows(" & i & ").RowHeight = " & Rows(i).Height, 1
        Next
        For i = fromColumn To toColumn
            '�񕝐ݒ�
            .WriteText "    Columns(" & i & ").ColumnWidth = " & Columns(i).ColumnWidth, 1
        Next
        
        .WriteText "    Dim onShape As Object", 1
        For Each shp In ActiveSheet.Shapes
        
            If shp.AutoShapeType <> msoShapeMixed Then
                .WriteText "    Set onShape = ActiveSheet.Shapes.AddShape(" & shp.AutoShapeType & "," & shp.Left & "," & shp.Top & "," & shp.Width & "," & shp.Height & ")", 1
                .WriteText "    onShape.Name = """ & shp.Name & """", 1
                .WriteText "    onShape.Visible = " & shp.Visible, 1
                .WriteText "    onShape.Line.ForeColor.RGB = " & shp.Line.ForeColor.RGB, 1
                .WriteText "    onShape.ForeColor.RGB = " & shp.ForeColor.RGB, 1
                '.WriteText "    onShape.TextFrame2.WordArtformat = " & shp.TextFrame2.WordArtformat, 1
                .WriteText "    onShape.Fill.Transparency = " & shp.Fill.Transparency, 1
                .WriteText "    onShape.TextFrame.Characters.Text = """ & shp.TextFrame.Characters.Text & """", 1
                .WriteText "    onShape.Fill.ForeColor.RGB = " & shp.Fill.ForeColor.RGB, 1
                .WriteText "    onShape.TextFrame2.TextRange.Font.Size = " & shp.TextFrame2.TextRange.Font.Size, 1
                .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
                .WriteText "    onShape.TextFrame.Characters.Font.Color = " & shp.TextFrame.Characters.Font.color, 1
                .WriteText "    onShape.TextFrame.Characters.Font.Name = """ & shp.TextFrame.Characters.Font.Name & """", 1
                .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
                .WriteText "    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = " & shp.TextFrame2.TextRange.ParagraphFormat.Alignment, 1
                .WriteText "    onShape.Placement = " & shp.Placement, 1
                .WriteText "    onShape.LockAspectRatio = " & shp.LockAspectRatio, 1
                .WriteText "    onShape.TextFrame2.AutoSize = " & shp.TextFrame2.AutoSize, 1
                .WriteText "    onShape.TextFrame2.MarginLeft = " & shp.TextFrame2.MarginLeft, 1
                .WriteText "    onShape.TextFrame2.MarginRight = " & shp.TextFrame2.MarginRight, 1
                .WriteText "    onShape.TextFrame2.MarginTop = " & shp.TextFrame2.MarginTop, 1
                .WriteText "    onShape.TextFrame2.MarginBottom = " & shp.TextFrame2.MarginBottom, 1
                .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
                .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
                .WriteText "    onShape.TextFrame2.HorizontalAnchor = " & shp.TextFrame2.HorizontalAnchor, 1
                .WriteText "    onShape.TextFrame2.Orientation = " & shp.TextFrame2.Orientation, 1
                
            ElseIf shp.Connector Then
                .WriteText "    Set onShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorElbow,BeginX:=0,BeginY:=0,EndX:=0,EndY:=0)", 1
                .WriteText "    onShape.ConnectorFormat.Type = " & shp.ConnectorFormat.Type, 1
                .WriteText "    onShape.Name = """ & shp.Name & """", 1
                .WriteText "    onShape.Line.ForeColor.RGB = " & shp.Line.ForeColor.RGB, 1
                .WriteText "    onShape.Placement = " & shp.Placement, 1
                .WriteText "    onShape.LockAspectRatio = " & shp.LockAspectRatio, 1
                .WriteText "    onShape.top = " & shp.Top, 1
                .WriteText "    onShape.left = " & shp.Left, 1
                .WriteText "    onShape.width = " & shp.Width, 1
                .WriteText "    onShape.height = " & shp.Height, 1
                .WriteText "    onShape.Line.BeginArrowheadStyle = " & shp.Line.BeginArrowheadStyle, 1
                .WriteText "    onShape.Line.EndArrowheadStyle = " & shp.Line.EndArrowheadStyle, 1
                .WriteText "    onShape.Line.Weight = " & shp.Line.weight, 1

            ElseIf shp.Type = msoFormControl Then
                .WriteText "    Set onShape = ActiveSheet.Buttons.Add(" & shp.Left & "," & shp.Top & "," & shp.Width & "," & shp.Height & ")", 1
                .WriteText "    onShape.OnAction = """ & Mid(shp.OnAction, InStr(shp.OnAction, "!") + 1) & """", 1
                .WriteText "    onShape.Name = """ & shp.Name & """", 1
                .WriteText "    onShape.Visible = " & shp.Visible, 1
                .WriteText "    onShape.Placement = " & shp.Placement, 1
                .WriteText "    onShape.Characters.Text = """"", 1
                '�R�����g�A�E�g���͂Ȃ����o�͂���Ȃ�
                .WriteText "    onShape.Characters.Text = """ & shp.Characters.Text & """", 1
'                .WriteText "    onShape.Text = """ & shp.Text & """", 1
'                .WriteText "    onShape.Caption = """ & shp.Caption & """", 1
'                .WriteText "    onShape.TextFrame2.Characters.Text = """ & shp.TextFrame2.Characters.Text & """", 1
'                .WriteText "    onShape.TextFrame2.TextRange.Font.Size = " & shp.TextFrame2.TextRange.Font.Size, 1
'                .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
'                .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
'                .WriteText "    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = " & shp.TextFrame2.TextRange.ParagraphFormat.Alignment, 1
'                .WriteText "    onShape.TextFrame2.AutoSize = " & shp.TextFrame2.AutoSize, 1
'                .WriteText "    onShape.TextFrame2.MarginLeft = " & shp.TextFrame2.MarginLeft, 1
'                .WriteText "    onShape.TextFrame2.MarginRight = " & shp.TextFrame2.MarginRight, 1
'                .WriteText "    onShape.TextFrame2.MarginTop = " & shp.TextFrame2.MarginTop, 1
'                .WriteText "    onShape.TextFrame2.MarginBottom = " & shp.TextFrame2.MarginBottom, 1
'                .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
'                .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
'                .WriteText "    onShape.TextFrame2.HorizontalAnchor = " & shp.TextFrame2.HorizontalAnchor, 1
'                .WriteText "    onShape.TextFrame2.Orientation = " & shp.TextFrame2.Orientation, 1
            
            End If
            '�֐���؂�i���s���ɋN����u�v���V�[�W�����傫�����܂��B�v��������邽�߁j
            If count > 30 And count Mod 30 = 0 Then
                .WriteText "end Function", 1
                .WriteText "Function revival" & functionCount & "()", 1
                functionCount = functionCount + 1
            End If
            count = count + 1
                
        Next
        '�Ō��1�s�͉��s�Ȃ�
        .WriteText "End Function", 1
        
        .WriteText "Sub revival()", 1
        If IsNumeric(Right(moduleName, 1)) Then
            '���W���[�����̍Ōオ���l�̏ꍇ�A�V�[�g�쐬������ǉ�����
            .WriteText "    Worksheets.Add After:=Worksheets(Worksheets.Count)", 1
        End If
        
        For i = 0 To functionCount - 1
            .WriteText "    CALL revival" & i & "()", 1
        Next
        
        If IsNumeric(Right(moduleName, 1)) Then
            .WriteText "    Worksheets(1).select", 1
        End If
        .WriteText "end sub", 0
        
        If CHAR_SET = "UTF-8" Then
            'Stream�I�u�W�F�N�g�̐擪����̈ʒu���w�肷��BType�ɒl��ݒ肷��Ƃ���0�ł���K�v������
            .Position = 0
            '�����f�[�^��ނ��o�C�i���f�[�^�ɕύX����
            .Type = 1
            '�ǂݎ��J�n�ʒu�H��3�o�C�g�ڂɈړ�����i3�o�C�g��BOM�t���������폜���邽�߁j
            .Position = 3
            '�o�C�g�������ꎞ�ۑ�
            bytetmp = .read
            '�����ł͕ۑ��͕s�v�B��x���ď������񂾓��e�����Z�b�g����ړI������
            .Close
            '�ēx�J����
            .Open
            '�o�C�g�`���ŏ�������
            .write bytetmp
        End If
        '�ۑ�
        .SaveToFile filePath, 2
        '�R�s�[��t�@�C�������
        .Close
    End With
End Function
