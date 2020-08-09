ans = MsgBox("スタイル数が多いと20分くらいかかるよ。" & vbCr & "その間エクセル操作すると強制終了しちゃうよ！", vbYesNo, "ビルトイン以外のスタイルを削除します。いいですか？")

If ans = vbNo Then WScript.Quit

MsgBox "処理開始"

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

For Each strFname In WScript.Arguments
    Set objDoc = objExcel.Workbooks.Open(strFname)

    '全体スタイル数
    All = objDoc.Styles.Count
    '進捗
    i = 0

    'スタイル削除
    For Each S In objDoc.Styles
        '進捗表示
        objExcel.StatusBar = CStr(i) & " / " & CStr(All)
        'ビルトインは除く
        If Not S.BuiltIn Then
            S.Delete
        End If

        i = i + 1
    Next

    '進捗表示を消す
    objExcel.StatusBar = False
    '保存＆クローズ
    objDoc.Save
    objDoc.Close

    Set objDoc = Nothing
Next

objExcel.Quit