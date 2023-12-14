Function maximoDosVetores(vetor As Variant) As Variant
    Dim numeroVetores As Double
    Dim numeroElementos As Double
    Dim valorMaximo As Double
    Dim valorResultante(100) As Double
    
    numeroVetores = UBound(vetor)
    numeroElementos = UBound(vetor(0))
    
    For i = 0 To numeroElementos
        valorResultante(i) = 0
        For j = 0 To numeroVetores
            If vetor(j)(i) > valorResultante(i) Then
                valorResultante(i) = vetor(j)(i)
            End If
        Next j
    Next i
      
    maximoDosVetores = valorResultante
    
End Function
