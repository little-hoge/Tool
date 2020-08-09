'Excelファイルの全シートの表示体裁を整える
'    (1)カーソルを先頭(セルA1)へ
'    (2)ブックの表示を「標準」へ
'    (3)ズーム(表示倍率)を100%へ
'excel_force_normalview.vbs
 
'引数のチェック
Set args = WScript.Arguments
If args.Count > 0 Then
    file = args(0)
    If Right(file,4) = ".xls" Or Right(file,5) = ".xlsx" Then
        'Excelファイルは続行する
    Else
        Msgbox "Excelファイルではありません。", vbExclamation, "警告"
        WScript.Quit
    End If
Else
    Msgbox "Excelファイルをvbsファイルにドラッグアンドドロップしてください。", vbInformation, "情報"
    WScript.Quit
End If
 
'Excelオブジェクトを取得
Set excelApp = CreateObject("Excel.Application")
'ブックを開く
Set workbook = excelApp.Workbooks.Open(file)
excelApp.Visible = True
 
'各シートの操作
For Each sheet In workbook.Worksheets
    sheet.Activate
    'カーソルを先頭へ
    If sheet.Visible Then
        excelApp.Goto sheet.Range("A1"), True
    End If
    '改ページプレビュー
    excelApp.ActiveWindow.View = 2 'xlPageBreakPreview
    excelApp.ActiveWindow.Zoom = 85
    '標準
    excelApp.ActiveWindow.View = 1 'xlNormalView
    excelApp.ActiveWindow.Zoom = 100
Next
 
'先頭シートを選択
workbook.Worksheets(1).Activate
 
'ブックを保存
workbook.Save
'ブックを閉じる
workbook.Close
Set workbook = Nothing
excelApp.Quit
 
Msgbox "処理が完了しました。" & vbNewLine & file, vbInformation, "情報"
WScript.Quit