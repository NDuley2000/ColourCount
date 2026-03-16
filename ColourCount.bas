Function CountByColour(rng As Range, refCell As Range) As Long
    Dim cell As Range
    Dim counter As Long
    Dim targetColor As Long
    
    targetColor = refCell.Interior.Color
    
    For Each cell In rng
        If Not cell.Interior.ColorIndex = xlNone Then
            If cell.Interior.Color = targetColor Then
                counter = counter + 1
            End If
        End If
    Next cell
    
    CountByColor = counter
End Function
