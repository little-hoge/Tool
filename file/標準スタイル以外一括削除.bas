Attribute VB_Name = "�W���X�^�C���ȊO�ꊇ�폜"
Public Sub �W���X�^�C���ȊO�ꊇ�폜()
    Dim s
    Dim i, cnt As Long
    On Error Resume Next
    
    cnt = ActiveWorkbook.Styles.Count
    i = 0
    For Each s In ActiveWorkbook.Styles
        DoEvents
        '�W���X�^�C���͏���
        If Not s.BuiltIn Then
            s.Delete
            Application.StatusBar = "�X�^�C���폜���F" & i & "/" & cnt
            i = i + 1
        End If
    Next
    Application.StatusBar = ""
   
   
    
End Sub
