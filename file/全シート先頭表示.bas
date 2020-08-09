Attribute VB_Name = "全シート先頭表示"
'
' 全てのシートを左上選択状態にする
'
Public Sub 全シート先頭表示()
    On Error Resume Next
    
    ' 画面更新の停止
    Application.ScreenUpdating = False

    ' 全シート分実行
    Dim i As Integer
    For i = 1 To Worksheets.Count
        ' シートをアクティブ化
        Worksheets(i).Activate
        
        ' 左上にスクロール
        Dim j As Integer
        For j = 1 To Windows(1).Panes.Count
            Windows(1).Panes(j).ScrollColumn = 1
            Windows(1).Panes(j).ScrollRow = 1
        Next
        
        ActiveSheet.Cells(1, 1).Select      ' 左上を選択
        ActiveWindow.View = xlNormalView    ' 改ページプレビュー解除
        
        ' シートによる為コメントアウト
'        ActiveWindow.Zoom = 100             ' 倍率100%
    Next
    
    ' 1番目のシートをアクティブ化
    Worksheets(1).Activate

    ' 画面更新の再開
    Application.ScreenUpdating = True
End Sub
