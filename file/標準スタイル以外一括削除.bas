Attribute VB_Name = "標準スタイル以外一括削除"
Public Sub 標準スタイル以外一括削除()
    Dim s
    Dim i, cnt As Long
    On Error Resume Next
    
    cnt = ActiveWorkbook.Styles.Count
    i = 0
    For Each s In ActiveWorkbook.Styles
        DoEvents
        '標準スタイルは除く
        If Not s.BuiltIn Then
            s.Delete
            Application.StatusBar = "スタイル削除中：" & i & "/" & cnt
            i = i + 1
        End If
    Next
    Application.StatusBar = ""
   
   
    
End Sub
