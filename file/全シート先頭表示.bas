Attribute VB_Name = "�S�V�[�g�擪�\��"
'
' �S�ẴV�[�g������I����Ԃɂ���
'
Public Sub �S�V�[�g�擪�\��()
    On Error Resume Next
    
    ' ��ʍX�V�̒�~
    Application.ScreenUpdating = False

    ' �S�V�[�g�����s
    Dim i As Integer
    For i = 1 To Worksheets.Count
        ' �V�[�g���A�N�e�B�u��
        Worksheets(i).Activate
        
        ' ����ɃX�N���[��
        Dim j As Integer
        For j = 1 To Windows(1).Panes.Count
            Windows(1).Panes(j).ScrollColumn = 1
            Windows(1).Panes(j).ScrollRow = 1
        Next
        
        ActiveSheet.Cells(1, 1).Select      ' �����I��
        ActiveWindow.View = xlNormalView    ' ���y�[�W�v���r���[����
        
        ' �V�[�g�ɂ��׃R�����g�A�E�g
'        ActiveWindow.Zoom = 100             ' �{��100%
    Next
    
    ' 1�Ԗڂ̃V�[�g���A�N�e�B�u��
    Worksheets(1).Activate

    ' ��ʍX�V�̍ĊJ
    Application.ScreenUpdating = True
End Sub
