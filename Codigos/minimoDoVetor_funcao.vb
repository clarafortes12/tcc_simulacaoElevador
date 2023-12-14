Function minimoDoVetor(minValor As Double, vetor As Variant) As Variant
    For i = 0 To UBound(vetor)
        If vetor(i) > minValor Then
            vetor(i) = minValor
        End If
    Next i

    minimoDoVetor = vetor
End Function
