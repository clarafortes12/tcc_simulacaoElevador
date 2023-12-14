Function trapmf(x As Variant, a As Double, b As Double, c As Double, d As Double) As Variant
    Dim y(100) As Double
    
    For intI = 0 To UBound(y)
        If x(intI) <= b Then
            If x(intI) > a And x(intI) < b Then
                y(intI) = (x(intI) - a) / (b - a)
            ElseIf x(intI) = b Then
                y(intI) = 1
            Else
                y(intI) = 0
            End If
        ElseIf x(intI) >= c Then
            If x(intI) > c And x(intI) < d Then
                y(intI) = (d - x(intI)) / (d - c)
            ElseIf x(intI) = c Then
                y(intI) = 1
            Else
                y(intI) = 0
            End If
        ElseIf x(intI) < a Then
            y(intI) = 0
        ElseIf x(intI) > d Then
            y(intI) = 0
        Else
            y(intI) = 1
        End If

    Next
    
    trapmf = y
End Function
