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
'End Sub
