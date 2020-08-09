'Excel�t�@�C���̑S�V�[�g�̕\���̍ق𐮂���
'    (1)�J�[�\����擪(�Z��A1)��
'    (2)�u�b�N�̕\�����u�W���v��
'    (3)�Y�[��(�\���{��)��100%��
'excel_force_normalview.vbs
 
'�����̃`�F�b�N
Set args = WScript.Arguments
If args.Count > 0 Then
    file = args(0)
    If Right(file,4) = ".xls" Or Right(file,5) = ".xlsx" Then
        'Excel�t�@�C���͑��s����
    Else
        Msgbox "Excel�t�@�C���ł͂���܂���B", vbExclamation, "�x��"
        WScript.Quit
    End If
Else
    Msgbox "Excel�t�@�C����vbs�t�@�C���Ƀh���b�O�A���h�h���b�v���Ă��������B", vbInformation, "���"
    WScript.Quit
End If
 
'Excel�I�u�W�F�N�g���擾
Set excelApp = CreateObject("Excel.Application")
'�u�b�N���J��
Set workbook = excelApp.Workbooks.Open(file)
excelApp.Visible = True
 
'�e�V�[�g�̑���
For Each sheet In workbook.Worksheets
    sheet.Activate
    '�J�[�\����擪��
    If sheet.Visible Then
        excelApp.Goto sheet.Range("A1"), True
    End If
    '���y�[�W�v���r���[
    excelApp.ActiveWindow.View = 2 'xlPageBreakPreview
    excelApp.ActiveWindow.Zoom = 85
    '�W��
    excelApp.ActiveWindow.View = 1 'xlNormalView
    excelApp.ActiveWindow.Zoom = 100
Next
 
'�擪�V�[�g��I��
workbook.Worksheets(1).Activate
 
'�u�b�N��ۑ�
workbook.Save
'�u�b�N�����
workbook.Close
Set workbook = Nothing
excelApp.Quit
 
Msgbox "�������������܂����B" & vbNewLine & file, vbInformation, "���"
WScript.Quit