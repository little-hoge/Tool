Attribute VB_Name = "結合セルを解除して値埋め"
' Ctrl+A→Alt→H→M→U→Ctrl+G→Alt+S→K→Enter→Shift+?→↑→Ctrl+Enter
Sub 結合セルを解除して値埋め()
    Dim rng As Range
  
        For Each rng In ActiveSheet.UsedRange
            If rng.MergeCells Then
                With rng.MergeArea
                    .UnMerge
                    .Value = .Resize(1, 1).Value
            End With
        End If
    Next
End Sub

