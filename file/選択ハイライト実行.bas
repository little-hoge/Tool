Attribute VB_Name = "�I���n�C���C�g���s"
Public Sub �N���[�Y�n�C���C�g�I�t()
    Dim exp1 As String
    Dim exp2 As String
    Dim preExp As String
    Dim cnt As Long
    Dim i As Long
    Dim flg As Boolean
    
    flg = False
    
    exp1 = "=CELL(""address"")=ADDRESS(ROW(),COLUMN())"
    exp2 = "=OR(CELL(""row"")=ROW(),CELL(""col"")=COLUMN())"
    cnt = Cells.FormatConditions.Count
    
    If cnt <> 0 Then
        For i = cnt To 1 Step -1
            preExp = ""
            
            On Error Resume Next
            preExp = Cells.FormatConditions(i).Formula1
            On Error GoTo 0
            If preExp = exp1 Or preExp = exp2 Then
                Call �n�C���C�g�I�t(i)
                flag = True
            End If
        Next i
    End If
    
End Sub

'//�I���n�C���C�g�����̃��C���v���V�[�W��
Public Sub �I���n�C���C�g���s()
    Dim exp1 As String       '//������1
    Dim exp2 As String       '//������2
    Dim preExp As String     '//�ݒ�Ϗ�����
    Dim cnt As Long          '//�ݒ肳��Ă�������t�������̐�
    Dim i As Long
    Dim flg As Boolean       '//�n�C���C�gOFF����������
    flg = False
    exp1 = "=CELL(""address"") = ADDRESS(ROW(), COLUMN())"
    exp2 = "=OR(CELL(""row"")=ROW(), CELL(""col"")=COLUMN())"
    cnt = Cells.FormatConditions.Count
    
     '//�����t�����������݂���Ȃ�n�C���C�g�ݒ�ς��𔻒�
    If cnt <> 0 Then
        For i = cnt To 1 Step -1
            preExp = ""
            '//�n�C���C�g�ݒ�ςȂ�n�C���C�gOFF
            On Error Resume Next
            preExp = Cells.FormatConditions(i).Formula1
            On Error GoTo 0
            If preExp = exp1 Or preExp = exp2 Then
                Call �n�C���C�g�I�t(i)
                flg = True
            End If
        Next i
    End If
    
    '//�n�C���C�g�ݒ肪����Ă��Ȃ���΃n�C���C�gON
    If flg = False Then
        Call �n�C���C�g�I��
    End If
    
End Sub

'//�n�C���C�g������ON
Private Sub �n�C���C�g�I��()
    Dim fc As FormatCondition       '//�����t������
    Dim exp As String               '//����
    
    '//�I���Z���͔w�i�F�Ȃ�
    exp = "=CELL(""address"") = ADDRESS(ROW(), COLUMN())"
    Set fc = Cells.FormatConditions.Add(xlExpression, , exp)
    fc.Interior.ColorIndex = 0
    
    '//�I���Z���Ɠ���A���s���n�C���C�g
    exp = "=OR(CELL(""row"")=ROW(), CELL(""col"")=COLUMN())"
    Set fc = Cells.FormatConditions.Add(xlExpression, , exp)
    fc.Interior.ColorIndex = 44
    
End Sub

'//�n�C���C�g����OFF
Private Sub �n�C���C�g�I�t(i As Long)
    Cells.FormatConditions(i).Delete
End Sub
