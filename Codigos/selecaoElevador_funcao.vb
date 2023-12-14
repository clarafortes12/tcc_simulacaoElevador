Function selecaoElevador(distanciaElevador1Normalizada As Double, distanciaElevador2Normalizada As Double, tempoEsperaElevador1 As Double, tempoEsperaElevador2 As Double, cargaElevador1Normalizada As Double, cargaElevador2Normalizada As Double) As Double
    Dim elevador1 As Double
    Dim elevador2 As Double
    Dim elevadorSelecao As Integer
    Dim prioridadeMaior As Double
    
    If (cargaElevador1Normalizada >= 1 And cargaElevador2Normalizada < 1) Then
        elevadorSelecao = 2
    ElseIf (cargaElevador1Normalizada < 1 And cargaElevador2Normalizada >= 1) Then
        elevadorSelecao = 1
    Else
        If cargaElevador1Normalizada > 1 Then
            cargaElevador1Normalizada = 1
        End If
        
        If cargaElevador2Normalizada > 1 Then
            cargaElevador2Normalizada = 1
        End If
        
        If tempoEsperaElevador1 > 1 Then
            tempoEsperaElevador1 = 1
        End If
        
        If tempoEsperaElevador2 > 1 Then
            tempoEsperaElevador2 = 1
        End If
        
        If distanciaElevador1Normalizada > 1 Then
            distanciaElevador1Normalizada = 1
        End If
        
        If distanciaElevador2Normalizada > 1 Then
            distanciaElevador2Normalizada = 1
        End If
        
        elevador1 = resultadoFuzzi(distanciaElevador1Normalizada, tempoEsperaElevador1, cargaElevador1Normalizada)
        elevador2 = resultadoFuzzi(distanciaElevador2Normalizada, tempoEsperaElevador2, cargaElevador2Normalizada)
    
        prioridadeMaior = maximoEntreDoisNumeros(elevador1, elevador2)
        
        If prioridadeMaior = defuzzificacaoElevador1 Then
            elevadorSelecao = 1
        Else
            elevadorSelecao = 2
        End If
    End If
    
    selecaoElevador = elevadorSelecao

End Function
