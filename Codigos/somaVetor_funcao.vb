Function somaVetor(vetor As Variant) As Double
    Dim soma As Double
    soma = 0
    
    For i = 0 To UBound(vetor)
        soma = soma + vetor(i)
    Next i

    somaVetor = soma
End Function
