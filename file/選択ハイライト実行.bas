Attribute VB_Name = "選択ハイライト実行"
Public Sub クローズハイライトオフ()
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
                Call ハイライトオフ(i)
                flag = True
            End If
        Next i
    End If
    
End Sub

'//選択ハイライト処理のメインプロシージャ
Public Sub 選択ハイライト実行()
    Dim exp1 As String       '//条件式1
    Dim exp2 As String       '//条件式2
    Dim preExp As String     '//設定済条件式
    Dim cnt As Long          '//設定されている条件付き書式の数
    Dim i As Long
    Dim flg As Boolean       '//ハイライトOFFしたか判定
    flg = False
    exp1 = "=CELL(""address"") = ADDRESS(ROW(), COLUMN())"
    exp2 = "=OR(CELL(""row"")=ROW(), CELL(""col"")=COLUMN())"
    cnt = Cells.FormatConditions.Count
    
     '//条件付き書式が存在するならハイライト設定済かを判定
    If cnt <> 0 Then
        For i = cnt To 1 Step -1
            preExp = ""
            '//ハイライト設定済ならハイライトOFF
            On Error Resume Next
            preExp = Cells.FormatConditions(i).Formula1
            On Error GoTo 0
            If preExp = exp1 Or preExp = exp2 Then
                Call ハイライトオフ(i)
                flg = True
            End If
        Next i
    End If
    
    '//ハイライト設定がされていなければハイライトON
    If flg = False Then
        Call ハイライトオン
    End If
    
End Sub

'//ハイライト処理をON
Private Sub ハイライトオン()
    Dim fc As FormatCondition       '//条件付き書式
    Dim exp As String               '//数式
    
    '//選択セルは背景色なし
    exp = "=CELL(""address"") = ADDRESS(ROW(), COLUMN())"
    Set fc = Cells.FormatConditions.Add(xlExpression, , exp)
    fc.Interior.ColorIndex = 0
    
    '//選択セルと同列、同行をハイライト
    exp = "=OR(CELL(""row"")=ROW(), CELL(""col"")=COLUMN())"
    Set fc = Cells.FormatConditions.Add(xlExpression, , exp)
    fc.Interior.ColorIndex = 44
    
End Sub

'//ハイライト処理OFF
Private Sub ハイライトオフ(i As Long)
    Cells.FormatConditions(i).Delete
End Sub
