Attribute VB_Name = "Zππ΅Δlί"
' Ctrl+A¨Alt¨H¨M¨U¨Ctrl+G¨Alt+S¨K¨Enter¨Shift+?¨ͺ¨Ctrl+Enter
Sub Zππ΅Δlί()
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

