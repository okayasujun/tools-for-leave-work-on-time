Attribute VB_Name = "dev"
'#��Ɨp�̃V�F�C�v����i�J�����ɂ̂ݎg�p����j
'Sub �J���p_�V�F�C�v��1�����ɐ��񂳂���()
'    Dim top As Integer: top = Selection.top
'    Dim left As Integer: left = Selection.left
'
'    For Each moveShape In ActiveSheet.Shapes
'        moveShape.top = top + 20
'        moveShape.left = left + 20
'        top = moveShape.top
'        left = moveShape.left
'    Next
'End Sub
'#�Z�b�g�A�b�v�Ŏg�p����
'�g�p���̐��̐F�𒲂ׂ�v���V�[�W���B����V�F�C�v��������I��������ԂŎ��s���A�C�~�f�B�G�C�g�E�B���h�E���Q��
'Sub �J���p_�F�𒲂ׂ�()
'    Dim currentColorCode As Long: currentColorCode = Selection.ShapeRange.Item(1).Line.ForeColor.RGB
'    Dim Red As Integer: Red = currentColorCode Mod 256
'    Dim Green As Integer: Green = Int(currentColorCode / 256) Mod 256
'    Dim Blue As Integer: Blue = Int(currentColorCode / 256 / 256)
'
'    Debug.Print "�F�l�F" & currentColorCode
'    Debug.Print "�ԁF" & Red
'    Debug.Print "�΁F" & Green
'    Debug.Print "�F" & Blue
'    Debug.Print "RGB(" & Red & "," & Green; "," & Blue & ")"
'End Sub
'�h��Ԃ��̐F�𒲂ׂ����ꍇ�ͤ��L�\�[�X��2�s�ڂ��ȉ��ɕύX�����OK�
'    Dim currentColorCode As Long: currentColorCode = Selection.ShapeRange.Item(1).Fill.ForeColor.RGB
'�Z���̔w�i�F�𒲂ׂ����ꍇ�ͤ��L�\�[�X��2�s�ڂ��ȉ��ɕύX�����OK�
'    Dim currentColorCode As Long: currentColorCode = Selection.Interior.Color
'�Z���̕����F�𒲂ׂ����ꍇ�ͤ��L�\�[�X��2�s�ڂ��ȉ��ɕύX�����OK�
'    Dim currentColorCode As Long: currentColorCode = Selection.Font.Color
'Sub �J���p_���ׂẴR�l�N�^�V�F�C�v���폜����()
'    For Each shp In ActiveSheet.Shapes
'        If shp.Connector Then
'            shp.Delete
'        End If
'    Next
''End Sub
'Sub �I�𒆃V�F�C�v���Č�����\�[�X�𐶐�����()
'
'    Dim shp As shape
'    Set shp = Selection.ShapeRange.Item(1)
'    Dim filePath As String: filePath = ActiveWorkbook.Path & "\source-" & Format(Now(), "yyyymmddhhnn") & ".txt"
'    Const CHAR_SET = "SHIFT-JIS"
'    With CreateObject("ADODB.Stream")
'        .Charset = CHAR_SET
'        'https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/lineseparatorsenum?view=sql-server-ver16
'        .LineSeparator = 10
'        .Open
'        .WriteText "    Dim onShape As Object", 1
'        If shp.Type = msoAutoShape Then
'            .WriteText "    Set onShape = ActiveSheet.Shapes.AddShape(" & shp.AutoShapeType & "," & shp.left & "," & shp.top & "," & shp.width & "," & shp.height & ")", 1
'            .WriteText "    onShape.Name = """ & shp.Name & """", 1
'            .WriteText "    onShape.Visible = " & shp.Visible, 1
'            .WriteText "    onShape.Line.ForeColor.RGB = " & shp.Line.ForeColor.RGB, 1
'            '.WriteText "    onShape.ForeColor.RGB = " & shp.ForeColor.RGB, 1
'            '.WriteText "    onShape.TextFrame2.WordArtformat = " & shp.TextFrame2.WordArtformat, 1
'            .WriteText "    onShape.Fill.Transparency = " & shp.Fill.Transparency, 1
'            .WriteText "    onShape.TextFrame.Characters.Text = """ & shp.TextFrame.Characters.Text & """", 1
'            .WriteText "    onShape.Fill.ForeColor.RGB = " & shp.Fill.ForeColor.RGB, 1
'            .WriteText "    onShape.TextFrame2.TextRange.Font.Size = " & shp.TextFrame2.TextRange.Font.Size, 1
'            .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
'            .WriteText "    onShape.TextFrame.Characters.Font.Color = " & shp.TextFrame.Characters.Font.Color, 1
'            .WriteText "    onShape.TextFrame.Characters.Font.Name = """ & shp.TextFrame.Characters.Font.Name & """", 1
'            .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
'            .WriteText "    onShape.TextFrame2.TextRange.ParagraphFormat.Alignment = " & shp.TextFrame2.TextRange.ParagraphFormat.Alignment, 1
'            .WriteText "    onShape.Placement = " & shp.Placement, 1
'            .WriteText "    onShape.LockAspectRatio = " & shp.LockAspectRatio, 1
'            .WriteText "    onShape.TextFrame2.AutoSize = " & shp.TextFrame2.AutoSize, 1
'            .WriteText "    onShape.TextFrame2.MarginLeft = " & shp.TextFrame2.MarginLeft, 1
'            .WriteText "    onShape.TextFrame2.MarginRight = " & shp.TextFrame2.MarginRight, 1
'            .WriteText "    onShape.TextFrame2.MarginTop = " & shp.TextFrame2.MarginTop, 1
'            .WriteText "    onShape.TextFrame2.MarginBottom = " & shp.TextFrame2.MarginBottom, 1
'            .WriteText "    onShape.TextFrame2.WordWrap = " & shp.TextFrame2.WordWrap, 1
'            .WriteText "    onShape.TextFrame2.VerticalAnchor = " & shp.TextFrame2.VerticalAnchor, 1
'            .WriteText "    onShape.TextFrame2.HorizontalAnchor = " & shp.TextFrame2.HorizontalAnchor, 1
'            .WriteText "    onShape.TextFrame2.Orientation = " & shp.TextFrame2.Orientation, 1
'
'        ElseIf shp.Connector Then
'            .WriteText "    Set onShape = ActiveSheet.Shapes.AddConnector(Type:=msoConnectorElbow,BeginX:=0,BeginY:=0,EndX:=0,EndY:=0)", 1
'            .WriteText "    onShape.ConnectorFormat.Type = " & shp.ConnectorFormat.Type, 1
'            .WriteText "    onShape.Name = """ & shp.Name & """", 1
'            .WriteText "    onShape.Line.ForeColor.RGB = " & shp.Line.ForeColor.RGB, 1
'            .WriteText "    onShape.Placement = " & shp.Placement, 1
'            .WriteText "    onShape.LockAspectRatio = " & shp.LockAspectRatio, 1
'            .WriteText "    onShape.top = " & shp.top, 1
'            .WriteText "    onShape.left = " & shp.left, 1
'            .WriteText "    onShape.width = " & shp.width, 1
'            .WriteText "    onShape.height = " & shp.height, 1
'            .WriteText "    onShape.Line.BeginArrowheadStyle = " & shp.Line.BeginArrowheadStyle, 1
'            .WriteText "    onShape.Line.EndArrowheadStyle = " & shp.Line.EndArrowheadStyle, 1
'            .WriteText "    onShape.Line.Weight = " & shp.Line.Weight, 1
'        End If
'        '�ۑ�
'        .SaveToFile filePath, 2
'        '�R�s�[��t�@�C�������
'        .Close
'    End With
'End Sub
''�o�͂����\�[�X��\�t���ē���m�F����p
'Sub �o�͊m�F()
'
'End Sub
