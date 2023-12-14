Function defuzzificacao(normalizacao As Variant, saida As Variant) As Double
    
    Dim valorDefuzzificacao As Double
    Dim somaMomentoArea As Double
    Dim somaArea As Double
    
    Dim x1 As Double
    Dim x2 As Double
    Dim y1 As Double
    Dim y2 As Double
    
    Dim momentoArea As Double
    Dim area As Double
    
    Dim epsilon As Double
    epsilon = EPS

    
    If somaVetor(saida) = 0 Then
        valorDefuzzificacao = 0
    Else
        somaMomentoArea = 0
        somaArea = 0
        
        If UBound(normalizacao) = 1 Then
            somaMomentoArea = normalizacao(0) * saida(0)
            somaArea = saida(0)
        Else
            'centroide
            For i = 1 To UBound(normalizacao)
                x1 = normalizacao(i - 1)
                x2 = normalizacao(i)
                y1 = saida(i - 1)
                y2 = saida(i)
                
                If ((y1 <> 0 And y2 <> 0) Or x1 <> x2) Then
                    If y1 = y2 Then
                        momentoArea = 0.5 * (x1 + x2)
                        area = (x2 - x1) * y1
                    ElseIf y1 = 0 And y2 <> 0 Then
                        momentoArea = ((2 / 3) * (x2 - x1)) + x1
                        area = 0.5 * (x2 - x1) * y2
                    ElseIf y2 = 0 And y1 <> 0 Then
                        momentoArea = ((1 / 3) * (x2 - x1)) + x1
                        area = 0.5 * (x2 - x1) * y1
                    Else
                        momentoArea = (((2 / 3) * (x2 - x1) * (y2 + 0.5 * y1)) / (y1 + y2)) + x1
                        area = 0.5 * (x2 - x1) * (y1 + y2)
                    End If
                    somaMomentoArea = somaMomentoArea + (momentoArea * area)
                    somaArea = somaArea + area
                End If
            Next i
        End If
        
        valorDefuzzificacao = somaMomentoArea / maximoEntreDoisNumeros(somaArea, epsilon)
    End If
    
    defuzzificacao = valorDefuzzificacao
    
End Function
