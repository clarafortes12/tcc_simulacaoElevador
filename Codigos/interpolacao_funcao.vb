Function interpolacao(x As Variant, y As Variant, ponto As Double) As Double
'f(x)=f(x0) + (((f(x1)-f(x0))*(x-x0))/(x1 - x0))
'f(x) interpolacao para o ponto x
'x0 e x1 pontos conhecidos próximos a x

    Dim x0 As Double
    Dim x1 As Double
    Dim y0 As Double
    Dim y1 As Double
    Dim resultadoInterp As Double

'Encontrar os pontos conhecidos mais próximos a x
    For i = 0 To UBound(x)
        If x(i) <= ponto Then
            x0 = x(i)
            y0 = y(i)
        End If
        If x(i) >= ponto Then
            x1 = x(i)
            y1 = y(i)
            Exit For
        End If
    Next i
    
    If x0 = x1 Then
        resultadoInterp = y0
    Else
        'Calcular o valor interpolado usando a fórmula de interpolação linear
        resultadoInterp = y0 + (((y1 - y0) * (ponto - x0)) / (x1 - x0))
    End If
    
    'Retornar o resultado interpolado
    interpolacao = resultadoInterp

End Function
