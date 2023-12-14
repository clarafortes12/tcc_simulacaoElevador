Function trimf(x As Variant, a As Double, b As Double, c As Double) As Variant
    Dim y(100) As Double
    
    For intI = 0 To UBound(y)
        If x(intI) > a And x(intI) < b Then
            y(intI) = (x(intI) - a) / (b - a)
        ElseIf x(intI) > b And x(intI) < c Then
            y(intI) = (c - x(intI)) / (c - b)
        ElseIf x(intI) = b Then
            y(intI) = 1
        Else
            y(intI) = 0
        End If

    Next
    
    trimf = y
End Function