Attribute VB_Name = "textCodeJudge"
'�����W���[���̃\�[�X�͈ȉ�URL����ہX�R�s�[���Ă��܂��B
'http://nonsoft.la.coocan.jp/SoftSample/SampleModJUDG2.html
'
'****************************************************************************
' �@�\��    : Module1.bas
' �@�\����  : �����R�[�h����
' ���l      :
' ���쌠    : Copyright(C) 2008 - 2022 �̂� All rights reserved
' ---------------------------------------------------------------------------
' �g�p����  : ���̃T�C�g�̓��e���g�p(���p/����/�]��/���S��)�������ʕ���s����
'           : �����Ɍ��J/�z�z����ꍇ�́A���̃T�C�g���Q�l�ɂ����|���L�q���Ă�
'           : �������B(��)WEB�y�[�W��ReadMe�Ƀ����N��\���Ă�������
' ---------------------------------------------------------------------------
'****************************************************************************
Private Const JUDGEFIX = 9999           '�����R�[�h���聓
Private Const JUDGESIZEMAX = 100000     '�����R�[�h����o�C�g��
Private Const BinaryByteWeight = 0.5    '�o�C�i���@�����R�[�h�̈�v�d��
Private Const SingleByteWeight = 1      '�P�o�C�g�@�����R�[�h�̈�v�d��
Private Const Multi_ByteWeight = 2      '�����o�C�g�����R�[�h�̈�v�d��
Private Enum JISMODE                    'JIS�R�[�h�̃��[�h
    ctrl = 0                            '����R�[�h
    asci = 1                            'ASCII
    roma = 2                            'JIS���[�}��
    kana = 3                            'JIS�J�i�i���p�J�i�j
    kanO = 4                            '��JIS���� (1978)
    kanN = 5                            '�VJIS���� (1983/1990)
    kanH = 6                            'JIS�⏕����
End Enum

'----�����R�[�h����
' �֐���    : JudgeCode
' �Ԃ�l    : ���茋�ʕ����R�[�h��
' ������    : bytCode : ���蕶���f�[�^
' �@�\����  : �����R�[�h�𔻒肷��
' ���l      :
Public Function judgeCode(ByRef bytCode() As Byte, Optional ByVal blnBin As Boolean = False) As String
    judgeCode = "SJIS"
    Dim dblSJIS As Double
    Dim dblJIS As Double
    Dim dblEUC As Double
    Dim dblUNI As Double
    Dim dblUTF7 As Double
    Dim dblUTF8 As Double
    Dim dblBIN As Double
    
    dblJIS = JudgeJIS(bytCode, False) ': Debug.Print "JIS :" & dblJIS
    If dblJIS >= JUDGEFIX Then judgeCode = "JIS": Exit Function
    
    dblUNI = JudgeUNI(bytCode, False) ': Debug.Print "UNI :" & dblUNI
    If dblUNI >= JUDGEFIX Then judgeCode = "UNICODE": Exit Function
    
    dblUTF8 = JudgeUTF8(bytCode, False) ': Debug.Print "UTF8:" & dblUTF8
    If dblUTF8 >= JUDGEFIX Then judgeCode = "UTF8": Exit Function

    dblUTF7 = JudgeUTF7(bytCode, False) ': Debug.Print "UTF7:" & dblUTF7
    If dblUTF7 >= JUDGEFIX Then judgeCode = "UTF7": Exit Function
    
    dblSJIS = JudgeSJIS(bytCode, False) ': Debug.Print "SJIS:" & dblSJIS
    If dblSJIS >= JUDGEFIX Then judgeCode = "SJIS": Exit Function
    
    dblEUC = JudgeEUC(bytCode, False) ': Debug.Print "EUC :" & dblEUC
    If dblEUC >= JUDGEFIX Then judgeCode = "EUC": Exit Function
    
    If blnBin Then
        dblBIN = JudgeBIN(bytCode, False) ': Debug.Print "BIN :" & dblBIN
        If dblBIN >= JUDGEFIX Then judgeCode = "BIN": Exit Function
    Else
        dblBIN = 0
    End If
    'Debug.Print "--------"

    If dblSJIS >= dblSJIS And dblSJIS >= dblUNI And dblSJIS >= dblJIS And _
       dblSJIS >= dblUTF7 And dblSJIS >= dblUTF8 And dblSJIS >= dblEUC And _
       dblSJIS >= dblBIN Then
        judgeCode = "SJIS"
        Exit Function
    End If

    If dblUNI >= dblSJIS And dblUNI >= dblUNI And dblUNI >= dblJIS And _
       dblUNI >= dblUTF7 And dblUNI >= dblUTF8 And dblUNI >= dblEUC And _
       dblUNI >= dblBIN Then
        judgeCode = "UNICODE"
        Exit Function
    End If

    If dblJIS >= dblSJIS And dblJIS >= dblUNI And dblJIS >= dblJIS And _
       dblJIS >= dblUTF7 And dblJIS >= dblUTF8 And dblJIS >= dblEUC And _
       dblJIS >= dblBIN Then
        judgeCode = "JIS"
        Exit Function
    End If

    If dblUTF7 >= dblSJIS And dblUTF7 >= dblUNI And dblUTF7 >= dblJIS And _
       dblUTF7 >= dblUTF7 And dblUTF7 >= dblUTF8 And dblUTF7 >= dblEUC And _
       dblUTF7 >= dblBIN Then
        judgeCode = "UTF7"
        Exit Function
    End If

    If dblUTF8 >= dblSJIS And dblUTF8 >= dblUNI And dblUTF8 >= dblJIS And _
       dblUTF8 >= dblUTF7 And dblUTF8 >= dblUTF8 And dblUTF8 >= dblEUC And _
       dblUTF8 >= dblBIN Then
        judgeCode = "UTF8"
        Exit Function
    End If

    If dblEUC >= dblSJIS And dblEUC >= dblUNI And dblEUC >= dblJIS And _
       dblEUC >= dblUTF7 And dblEUC >= dblUTF8 And dblEUC >= dblEUC And _
       dblEUC >= dblBIN Then
        judgeCode = "EUC"
        Exit Function
    End If

    If dblBIN >= dblSJIS And dblBIN >= dblUNI And dblBIN >= dblJIS And _
       dblBIN >= dblUTF7 And dblBIN >= dblUTF8 And dblBIN >= dblEUC And _
       dblBIN >= dblBIN Then
        judgeCode = "BIN"
        Exit Function
    End If
    
End Function

'----SJIS�֌W
' �֐���    : JudgeSJIS
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : ���蕶���f�[�^�̔���m�����v�Z����
' ���l      :
Private Function JudgeSJIS(ByRef bytCode() As Byte, _
Optional ByVal fixFlag As Boolean = False) As Double
    Dim i As Long
    Dim dblFit As Double
    Dim dblUB As Double
    
    dblUB = JUDGESIZEMAX - 1
    If dblUB > UBound(bytCode) Then
        dblUB = UBound(bytCode)
    End If
    For i = 0 To dblUB
        '81-9F,E0-EF(1�o�C�g��)
        If (bytCode(i) >= &H81 And bytCode(i) <= &H9F) Or _
           (bytCode(i) >= &HE0 And bytCode(i) <= &HEF) Then
            If i <= UBound(bytCode) - 1 Then
                '40-7E,80-FC(2�o�C�g��)
                If (bytCode(i + 1) >= &H40 And bytCode(i + 1) <= &H7E) Or _
                   (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HFC) Then
                    dblFit = dblFit + (2 * Multi_ByteWeight)
                Else
                    dblFit = dblFit - (2 * Multi_ByteWeight)
                End If
                i = i + 1
            End If
        
        'A1-DF(1�o�C�g��)
        ElseIf (bytCode(i) >= &HA1 And bytCode(i) <= &HDF) Then
            dblFit = dblFit + (1 * SingleByteWeight)
        
        '20-7E(1�o�C�g��)
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            dblFit = dblFit + (1 * SingleByteWeight)
        
        '00-1F, 7F(1�o�C�g��)
        ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                bytCode(i) = &H7F Then
            If bytCode(i) = &H9 Or bytCode(i) = &HD Or bytCode(i) = &HA Then
                dblFit = dblFit + (1 * SingleByteWeight)
            Else
                dblFit = dblFit + (1 * BinaryByteWeight)
            End If
        End If
    Next i
    JudgeSJIS = (dblFit * 100) / ((dblUB + 1) * Multi_ByteWeight)
End Function

'----JIS�֌W
' �֐���    : JudgeJIS
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : ���蕶���f�[�^�̔���m�����v�Z����
' ���l      :
Private Function JudgeJIS(ByRef bytCode() As Byte, _
Optional ByVal fixFlag As Boolean = False) As Double
    Dim i As Long
    Dim dblFit As Double
    Dim dblUB As Double
    Dim lngMode As JISMODE
    
    dblUB = JUDGESIZEMAX - 1
    If dblUB > UBound(bytCode) Then
        dblUB = UBound(bytCode)
    End If
    For i = 0 To dblUB
        '1B(1�o�C�g��)
        If bytCode(i) = &H1B Then
            If i <= UBound(bytCode) - 2 Then
                '28 42(2�E3�o�C�g��)
                If bytCode(i + 1) = &H28 And bytCode(i + 1) <= &H42 Then
                    lngMode = asci
                    dblFit = dblFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    'If fixFlag Then
                    '    JudgeJIS = JUDGEFIX
                    '    Exit Function
                    'End If
                
                '28 4A(2�E3�o�C�g��)
                ElseIf bytCode(i + 1) = &H28 And bytCode(i + 1) <= &H4A Then
                    lngMode = roma
                    dblFit = dblFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    'If fixFlag Then
                    '    JudgeJIS = JUDGEFIX
                    '    Exit Function
                    'End If
                
                '28 49(2�E3�o�C�g��)
                ElseIf bytCode(i + 1) = &H28 And bytCode(i + 1) <= &H49 Then
                    lngMode = kana
                    dblFit = dblFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    'If fixFlag Then
                    '    JudgeJIS = JUDGEFIX
                    '    Exit Function
                    'End If
                
                '24 40(2�E3�o�C�g��)
                ElseIf bytCode(i + 1) = &H24 And bytCode(i + 1) <= &H40 Then
                    lngMode = kanO
                    dblFit = dblFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    'If fixFlag Then
                    '    JudgeJIS = JUDGEFIX
                    '    Exit Function
                    'End If
                
                '24 42(2�E3�o�C�g��)
                ElseIf bytCode(i + 1) = &H24 And bytCode(i + 1) <= &H42 Then
                    lngMode = kanN
                    dblFit = dblFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    'If fixFlag Then
                    '    JudgeJIS = JUDGEFIX
                    '    Exit Function
                    'End If
                
                '24 44(2�E3�o�C�g��)
                ElseIf bytCode(i + 1) = &H24 And bytCode(i + 1) <= &H44 Then
                    lngMode = kanH
                    dblFit = dblFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    'If fixFlag Then
                    '    JudgeJIS = JUDGEFIX
                    '    Exit Function
                    'End If
                End If
            End If
        Else
            Select Case lngMode
            Case ctrl, asci, roma
                '00-1F,7F
                If (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                    bytCode(i) = &H7F Then
                    If bytCode(i) = &H9 Or bytCode(i) = &HD Or bytCode(i) = &HA Then
                        dblFit = dblFit + (1 * SingleByteWeight)
                    Else
                        dblFit = dblFit + (1 * BinaryByteWeight)
                    End If
                
                '20-7E
                ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
                    dblFit = dblFit + (1 * SingleByteWeight)
                End If
            Case kana
                '21-5F
                If (bytCode(i) >= &H21 And bytCode(i) <= &H5F) Then
                    dblFit = dblFit + (1 * SingleByteWeight)
                End If
            Case kanO, kanN, kanH
                If i <= UBound(bytCode) - 1 Then
                    '21-7E
                    If (bytCode(i) >= &H21 And bytCode(i) <= &H7E) And _
                       (bytCode(i - 1) >= &H21 And bytCode(i - 1) <= &H7E) Then
                        dblFit = dblFit + (2 * Multi_ByteWeight)
                        i = i + 1
                    End If
                End If
            End Select
        End If
    Next i
    JudgeJIS = (dblFit * 100) / ((dblUB + 1) * Multi_ByteWeight)
End Function

'----EUC�֌W
' �֐���    : JudgeEUC
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : ���蕶���f�[�^�̔���m�����v�Z����
' ���l      :
Private Function JudgeEUC(ByRef bytCode() As Byte, _
Optional ByVal fixFlag As Boolean = False) As Double
    Dim i As Long
    Dim dblFit As Double
    Dim dblUB As Double
    
    dblUB = JUDGESIZEMAX - 1
    If dblUB > UBound(bytCode) Then
        dblUB = UBound(bytCode)
    End If
    For i = 0 To dblUB
        '8E(1�o�C�g��) + A1-DF(2�o�C�g��)
        If bytCode(i) = &H8E Then
            If i <= UBound(bytCode) - 1 Then
                If bytCode(i + 1) >= &HA1 And bytCode(i + 1) <= &HDF Then
                    dblFit = dblFit + (2 * Multi_ByteWeight)
                Else
                    dblFit = dblFit - (2 * Multi_ByteWeight)
                End If
                i = i + 1
            End If
        
        '8F(1�o�C�g��) + A1-0xFE(2�E3�o�C�g��)
        ElseIf bytCode(i) = &H8F Then
            If i <= UBound(bytCode) - 2 Then
                If (bytCode(i + 1) >= &HA1 And bytCode(i + 1) <= &HFE) And _
                   (bytCode(i + 2) >= &HA1 And bytCode(i + 2) <= &HFE) Then
                    dblFit = dblFit + (3 * Multi_ByteWeight)
                Else
                    dblFit = dblFit - (3 * Multi_ByteWeight)
                End If
                i = i + 2
            End If
        
        'A1-FE(1�o�C�g��) + A1-FE(2�o�C�g��)
        ElseIf bytCode(i) >= &HA1 And bytCode(i) <= &HFE Then
            If i <= UBound(bytCode) - 1 Then
                If bytCode(i + 1) >= &HA1 And bytCode(i + 1) <= &HFE Then
                    dblFit = dblFit + (2 * Multi_ByteWeight)
                Else
                    dblFit = dblFit - (2 * Multi_ByteWeight)
                End If
                i = i + 1
            End If
        
        '20-7E(1�o�C�g��)
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            dblFit = dblFit + (1 * SingleByteWeight)
        
        '00-1F, 7F(1�o�C�g��)
        ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                bytCode(i) = &H7F Then
            If bytCode(i) = &H9 Or bytCode(i) = &HD Or bytCode(i) = &HA Then
                dblFit = dblFit + (1 * SingleByteWeight)
            Else
                dblFit = dblFit + (1 * BinaryByteWeight)
            End If
        End If
    Next i
    JudgeEUC = (dblFit * 100) / ((dblUB + 1) * Multi_ByteWeight)
End Function

'----UNICODE�֌W
' �֐���    : JudgeUNI
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : ���蕶���f�[�^�̔���m�����v�Z����
' ���l      :
Private Function JudgeUNI(ByRef bytCode() As Byte, _
Optional ByVal fixFlag As Boolean = False) As Double
    Dim i As Long
    Dim dblFit As Double
    Dim dblUB As Double

    dblUB = JUDGESIZEMAX - 1
    If dblUB > UBound(bytCode) Then
        dblUB = UBound(bytCode)
    End If
    
    For i = 0 To dblUB
        If i = 0 And fixFlag Then
            'BOM
            If bytCode(i) = &HFF Then
                If i <= UBound(bytCode) - 1 Then
                    If bytCode(i + 1) = &HFE Then
                        JudgeUNI = JUDGEFIX
                        Exit Function
                    End If
                End If
            End If
            '���p�̏�
            'If bytCode(i) = &H0 Then
            '    JudgeUNI = JUDGEFIX
            '    Exit Function
            'End If
        End If
        
        If i <= UBound(bytCode) - 1 Then
            '00(2�o�C�g��)
            If (bytCode(i + 1) = &H0) Then
                '00-FF(1�o�C�g��)
                dblFit = dblFit + UniPoint(bytCode, i)
            
            '01-33(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &H1 And bytCode(i + 1) <= &H33) Then
                '00-FF(1�o�C�g��)
                dblFit = dblFit + UniPoint(bytCode, i)
            
            '34-4D(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &H34 And bytCode(i + 1) <= &H4D) Then
                '00-FF(1�o�C�g��)----��----
                dblFit = dblFit + 0
            
            '4E-9F(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &H4E And bytCode(i + 1) <= &H9F) Then
                '00-FF(1�o�C�g��)
                dblFit = dblFit + UniPoint(bytCode, i)
            
            'A0-AB(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HA0 And bytCode(i + 1) <= &HAB) Then
                '00-FF(1�o�C�g��)----��----
                dblFit = dblFit + 0
            
            'AC-D7(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HAC And bytCode(i + 1) <= &HD7) Then
                '00-FF(1�o�C�g��)----�n���O��----
                dblFit = dblFit + 0
            
            'D8-DF(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HD8 And bytCode(i + 1) <= &HDF) Then
                '00-FF(1�o�C�g��)
                dblFit = dblFit + UniPoint(bytCode, i)
            
            'E0-F7(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HE0 And bytCode(i + 1) <= &HF7) Then
                '00-FF(1�o�C�g��)----�O��----
                dblFit = dblFit + 0
            
            'F8-FF(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HF8 And bytCode(i + 1) <= &HFF) Then
                '00-FF(1�o�C�g��)
                dblFit = dblFit + UniPoint(bytCode, i)
            
            End If
            i = i + 1
        End If
    Next i
    JudgeUNI = (dblFit * 100) / ((dblUB + 1) * Multi_ByteWeight)
End Function
Private Function UniPoint(ByRef dat() As Byte, ByVal idx As Long) As Double
    On Error Resume Next
    UniPoint = 0
    If UBound(dat) <= idx - 1 Then Exit Function
    Dim ddd(1) As Byte
    ddd(0) = dat(idx)
    ddd(1) = dat(idx + 1)
    Dim eee As String
    'eee = System.Text.Encoding.GetEncoding("UNICODE").GetString(ddd)
    eee = ddd
    If eee = "" Then eee = "?"
    If Hex(Asc(eee)) <> "3F" Then
        If ddd(0) >= &H0 And ddd(0) <= &H7E And _
           ddd(1) >= &H1 And ddd(1) <= &H7E Then
            If ddd(0) <= &H1F Then
                UniPoint = UniPoint + BinaryByteWeight
            Else
                UniPoint = UniPoint + SingleByteWeight
            End If
            If ddd(1) <= &H1F Then
                UniPoint = UniPoint + BinaryByteWeight
            Else
                UniPoint = UniPoint + SingleByteWeight
            End If
        Else
            If ddd(1) = &H0 Then
                If ddd(0) <= &H1F Then
                    If ddd(0) = &H9 Or ddd(0) = &HD Or ddd(0) = &HA Then
                        UniPoint = UniPoint + (2 * Multi_ByteWeight)
                    Else
                        UniPoint = UniPoint + (2 * BinaryByteWeight)
                    End If
                Else
                    UniPoint = UniPoint + (2 * Multi_ByteWeight)
                End If
            Else
                UniPoint = UniPoint + (2 * Multi_ByteWeight)
            End If
        End If
    Else
    End If
End Function

'----UTF7�֌W
' �֐���    : JudgeUTF7
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : ���蕶���f�[�^�̔���m�����v�Z����
' ���l      :
Private Function JudgeUTF7(ByRef bytCode() As Byte, _
Optional ByVal fixFlag As Boolean = False) As Double
    Dim i As Long
    Dim dblFit As Double
    Dim dblUB As Double
    Dim lngWrk As Long
    Dim str64 As String
    Dim bln64 As Boolean
    str64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim lngBY As Long
    Dim lngXB As Long
    Dim lngXX As Long
    
    dblUB = JUDGESIZEMAX - 1
    If dblUB > UBound(bytCode) Then
        dblUB = UBound(bytCode)
    End If
    lngWrk = 0
    
    For i = 0 To dblUB
        '+�`-�܂ł�BASE64ENCODE
        If bytCode(i) = Asc("+") And bln64 = False Then
            lngWrk = 1
            bln64 = True
        ElseIf bytCode(i) = Asc("-") Then
            If lngWrk <= 0 Then
                lngWrk = lngWrk + 1
                dblFit = dblFit + (lngWrk * SingleByteWeight)
            ElseIf lngWrk = 1 Then
                lngWrk = lngWrk + 1
                dblFit = dblFit + (lngWrk * Multi_ByteWeight)
            ElseIf lngWrk >= 4 And lngXB < 6 And _
                   ((InStr(str64, Chr(bytCode(i - 1))) - 1) And lngXX) = 0 Then
                lngWrk = lngWrk + 1
                dblFit = dblFit + (lngWrk * Multi_ByteWeight)
            End If
            lngWrk = 0
            bln64 = False
        Else
            If bln64 = True Then
                'BASE64ENCODE��
                If InStr(str64, Chr(bytCode(i))) > 0 Then
                    lngBY = Int((lngWrk * 6) / 8)
                    lngXB = (lngWrk * 6) - (lngBY * 8)
                    lngXX = (2 ^ lngXB) - 1
                    lngWrk = lngWrk + 1
                Else
                    lngWrk = 0
                    bln64 = False
                End If
            Else
                '20-7E(1�o�C�g��)
                If (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
                    dblFit = dblFit + (1 * SingleByteWeight)
                
                '00-1F, 7F(1�o�C�g��)
                ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                        bytCode(i) = &H7F Then
                    If bytCode(i) = &H9 Or bytCode(i) = &HD Or bytCode(i) = &HA Then
                        dblFit = dblFit + (1 * SingleByteWeight)
                    Else
                        dblFit = dblFit + (1 * BinaryByteWeight)
                    End If
                End If
            End If
        End If
    Next i
    JudgeUTF7 = (dblFit * 100) / ((dblUB + 1) * Multi_ByteWeight)
End Function

'----UTF8�֌W
' �֐���    : JudgeUTF8
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : ���蕶���f�[�^�̔���m�����v�Z����
' ���l      :
Private Function JudgeUTF8(ByRef bytCode() As Byte, _
Optional ByVal fixFlag As Boolean = False) As Double
    Dim i As Long
    Dim dblFit As Double
    Dim dblUB As Double
    
    dblUB = JUDGESIZEMAX - 1
    If dblUB > UBound(bytCode) Then
        dblUB = UBound(bytCode)
    End If
    For i = 0 To dblUB
        If i = 0 And fixFlag Then
            'BOM
            If bytCode(i) = &HEF Then
                If i <= UBound(bytCode) - 2 Then
                    If bytCode(i + 1) = &HBB And _
                       bytCode(i + 2) = &HBF Then
                        JudgeUTF8 = JUDGEFIX
                        Exit Function
                    End If
                End If
            End If
        End If
        
        'AND FC(1�o�C�g��) + 80-BF(2-6�o�C�g��)
        If (bytCode(i) And &HFC) = &HFC Then
            If i <= UBound(bytCode) - 5 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) And _
                   (bytCode(i + 3) >= &H80 And bytCode(i + 3) <= &HBF) And _
                   (bytCode(i + 4) >= &H80 And bytCode(i + 4) <= &HBF) And _
                   (bytCode(i + 5) >= &H80 And bytCode(i + 5) <= &HBF) Then
                    dblFit = dblFit + (6 * Multi_ByteWeight)
                    i = i + 5
                End If
            End If
        
        'AND F8(1�o�C�g��) + 80-BF(2-5�o�C�g��)
        ElseIf (bytCode(i) And &HF8) = &HF8 Then
            If i <= UBound(bytCode) - 4 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) And _
                   (bytCode(i + 3) >= &H80 And bytCode(i + 3) <= &HBF) And _
                   (bytCode(i + 4) >= &H80 And bytCode(i + 4) <= &HBF) Then
                    dblFit = dblFit + (5 * Multi_ByteWeight)
                    i = i + 4
                End If
            End If
        
        'AND F0(1�o�C�g��) + 80-BF(2-4�o�C�g��)
        ElseIf (bytCode(i) And &HF0) = &HF0 Then
            If i <= UBound(bytCode) - 3 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) And _
                   (bytCode(i + 3) >= &H80 And bytCode(i + 3) <= &HBF) Then
                    dblFit = dblFit + (4 * Multi_ByteWeight)
                    i = i + 3
                End If
            End If
        
        'AND E0(1�o�C�g��) + 80-BF(2-3�o�C�g��)
        ElseIf (bytCode(i) And &HE0) = &HE0 Then
            If i <= UBound(bytCode) - 2 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) Then
                    dblFit = dblFit + (3 * Multi_ByteWeight)
                    i = i + 2
                End If
            End If
        
        'AND C0(1�o�C�g��) + 80-BF(2�o�C�g��)
        ElseIf (bytCode(i) And &HC0) = &HC0 Then
            If i <= UBound(bytCode) - 1 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) Then
                    dblFit = dblFit + (2 * Multi_ByteWeight)
                    i = i + 1
                End If
            End If
        
        '20-7E(1�o�C�g��)
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            dblFit = dblFit + (1 * SingleByteWeight)
        
        '00-1F, 7F(1�o�C�g��)
        ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                bytCode(i) = &H7F Then
            If bytCode(i) = &H9 Or bytCode(i) = &HD Or bytCode(i) = &HA Then
                dblFit = dblFit + (1 * SingleByteWeight)
            Else
                dblFit = dblFit + (1 * BinaryByteWeight)
            End If
        End If
    Next i
    JudgeUTF8 = (dblFit * 100) / ((dblUB + 1) * Multi_ByteWeight)
End Function

'----BIN�֌W
' �֐���    : JudgeBIN
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : ���蕶���f�[�^�̔���m�����v�Z����
' ���l      :
Private Function JudgeBIN(ByRef bytCode() As Byte, _
Optional ByVal fixFlag As Boolean = False) As Double
    Dim i As Long
    Dim dblFit As Double
    Dim dblUB As Double
    Dim intBin As Long

    dblUB = JUDGESIZEMAX - 1
    If dblUB > UBound(bytCode) Then
        dblUB = UBound(bytCode)
    End If
    For i = 0 To dblUB
        '00-1F, 7F
        If (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
            bytCode(i) = &H7F Then
            If bytCode(i) = &H9 Or bytCode(i) = &HD Or bytCode(i) = &HA Then
                dblFit = dblFit + (1 * SingleByteWeight)
                intBin = 0
            Else
                intBin = intBin + 1
                If intBin >= 2 Then
                    dblFit = dblFit + (1 * Multi_ByteWeight)
                Else
                    dblFit = dblFit + (1 * Multi_ByteWeight)
                End If
            End If

            '20-7E
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            dblFit = dblFit + (1 * SingleByteWeight)

        '80-FF
        Else
            dblFit = dblFit + (1 * SingleByteWeight)
            intBin = 0
        End If
    Next i
    JudgeBIN = (dblFit * 100) / ((dblUB + 1) * Multi_ByteWeight)
End Function

