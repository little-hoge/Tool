ans = MsgBox("�X�^�C������������20�����炢�������B" & vbCr & "���̊ԃG�N�Z�����삷��Ƌ����I�������Ⴄ��I", vbYesNo, "�r���g�C���ȊO�̃X�^�C�����폜���܂��B�����ł����H")

If ans = vbNo Then WScript.Quit

MsgBox "�����J�n"

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

For Each strFname In WScript.Arguments
    Set objDoc = objExcel.Workbooks.Open(strFname)

    '�S�̃X�^�C����
    All = objDoc.Styles.Count
    '�i��
    i = 0

    '�X�^�C���폜
    For Each S In objDoc.Styles
        '�i���\��
        objExcel.StatusBar = CStr(i) & " / " & CStr(All)
        '�r���g�C���͏���
        If Not S.BuiltIn Then
            S.Delete
        End If

        i = i + 1
    Next

    '�i���\��������
    objExcel.StatusBar = False
    '�ۑ����N���[�Y
    objDoc.Save
    objDoc.Close

    Set objDoc = Nothing
Next

objExcel.Quit