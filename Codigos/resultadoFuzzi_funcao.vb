Function resultadoFuzzi(distanciaElevador1Normalizada As Double, tempoEsperaElevador1 As Double, cargaElevador1Normalizada As Double) As Double
    Const tamanho As Integer = 100
    Dim normalizacao(tamanho) As Double
    
    Dim tempoEsperaBaixo As Variant
    Dim tempoEsperaMedio As Variant
    Dim tempoEsperaAlto As Variant
    Dim disponibilidadeCargaBaixo As Variant
    Dim disponibilidadeCargaMedio As Variant
    Dim disponibilidadeCargaAlto As Variant
    Dim distanciaBaixo As Variant
    Dim distanciaMedio As Variant
    Dim distanciaAlto As Variant
    Dim prioridadeBaixo As Variant
    Dim prioridadeMedio As Variant
    Dim prioridadeAlto As Variant
    
    Dim tempoEsperaBaixoElevador1 As Double
    Dim tempoEsperaMedioElevador1 As Double
    Dim tempoEsperaAltoElevador1 As Double
    Dim cargaBaixoElevador1 As Double
    Dim cargaMedioElevador1 As Double
    Dim cargaAltoElevador1 As Double
    Dim distanciaBaixoElevador1 As Double
    Dim distanciaMedioElevador1 As Double
    Dim distanciaAltoElevador1 As Double
    
    'Regras Elevador 1
    Dim rule1Elevador1 As Variant
    Dim rule2Elevador1 As Variant
    Dim rule3Elevador1 As Variant
    Dim rule4Elevador1 As Variant
    Dim rule5Elevador1 As Variant
    Dim rule6Elevador1 As Variant
    Dim rule7Elevador1 As Variant
    Dim rule8Elevador1 As Variant
    Dim rule9Elevador1 As Variant
    Dim rule10Elevador1 As Variant
    Dim rule11Elevador1 As Variant
    
    Dim saidaBaixaElevador1 As Variant
    Dim saidaMediaElevador1 As Variant
    Dim saidaAltaElevador1 As Variant
    
    Dim saidaFinalElevador1 As Variant
    Dim defuzzificacaoElevador1 As Double

    
    For intI = 0 To UBound(normalizacao)
        normalizacao(intI) = intI / UBound(normalizacao)
    Next
    
    tempoEsperaBaixo = trapmf(normalizacao, 0, 0, 0.3, 0.5)
    tempoEsperaMedio = trimf(normalizacao, 0.25, 0.55, 0.85)
    tempoEsperaAlto = trimf(normalizacao, 0.6, 1, 1)
    
    disponibilidadeCargaBaixo = trapmf(normalizacao, 0, 0, 0.3, 0.5)
    disponibilidadeCargaMedio = trimf(normalizacao, 0.25, 0.55, 0.85)
    disponibilidadeCargaAlto = trimf(normalizacao, 0.6, 1, 1)
    
    distanciaBaixo = trimf(normalizacao, 0, 0, 0.4)
    distanciaMedio = trimf(normalizacao, 0.2, 0.45, 0.8)
    distanciaAlto = trimf(normalizacao, 0.7, 1, 1)
    
    prioridadeBaixo = trimf(normalizacao, 0, 0, 0.4)
    prioridadeMedio = trimf(normalizacao, 0.1, 0.5, 0.9)
    prioridadeAlto = trimf(normalizacao, 0.6, 1, 1)

    tempoEsperaBaixoElevador1 = interpolacao(normalizacao, tempoEsperaBaixo, tempoEsperaElevador1)
    tempoEsperaMedioElevador1 = interpolacao(normalizacao, tempoEsperaMedio, tempoEsperaElevador1)
    tempoEsperaAltoElevador1 = interpolacao(normalizacao, tempoEsperaAlto, tempoEsperaElevador1)
    
    cargaBaixoElevador1 = interpolacao(normalizacao, disponibilidadeCargaBaixo, cargaElevador1Normalizada)
    cargaMedioElevador1 = interpolacao(normalizacao, disponibilidadeCargaMedio, cargaElevador1Normalizada)
    cargaAltoElevador1 = interpolacao(normalizacao, disponibilidadeCargaAlto, cargaElevador1Normalizada)
    
    distanciaBaixoElevador1 = interpolacao(normalizacao, distanciaBaixo, distanciaElevador1Normalizada)
    distanciaMedioElevador1 = interpolacao(normalizacao, distanciaMedio, distanciaElevador1Normalizada)
    distanciaAltoElevador1 = interpolacao(normalizacao, distanciaAlto, distanciaElevador1Normalizada)

    rule1Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(distanciaAltoElevador1, tempoEsperaAltoElevador1), prioridadeBaixo)
    rule2Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(distanciaAltoElevador1, tempoEsperaBaixoElevador1), prioridadeMedio)
    rule3Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(distanciaAltoElevador1, tempoEsperaMedioElevador1), prioridadeBaixo)
    rule4Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(distanciaBaixoElevador1, tempoEsperaAltoElevador1), prioridadeMedio)
    rule5Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(distanciaBaixoElevador1, tempoEsperaBaixoElevador1), prioridadeAlto)
    rule6Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(distanciaBaixoElevador1, tempoEsperaMedioElevador1), prioridadeMedio)
    rule7Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(distanciaMedioElevador1, tempoEsperaMedioElevador1), prioridadeMedio)
    rule8Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(distanciaMedioElevador1, tempoEsperaAltoElevador1), prioridadeBaixo)
    rule9Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(cargaBaixoElevador1, tempoEsperaAltoElevador1), prioridadeBaixo)
    rule10Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(distanciaBaixoElevador1, cargaBaixoElevador1), prioridadeBaixo)
    rule11Elevador1 = minimoDoVetor(minimoEntreDoisNumeros(cargaMedioElevador1, tempoEsperaBaixoElevador1), prioridadeAlto)
    
    saidaBaixaElevador1 = maximoDosVetores(Array(rule1Elevador1, rule3Elevador1, rule8Elevador1, rule9Elevador1, rule10Elevador1))
    saidaMediaElevador1 = maximoDosVetores(Array(rule2Elevador1, rule4Elevador1, rule6Elevador1, rule7Elevador1))
    saidaAltaElevador1 = maximoDosVetores(Array(rule5Elevador1, rule11Elevador1))

    saidaFinalElevador1 = maximoDosVetores(Array(saidaBaixaElevador1, saidaMediaElevador1, saidaAltaElevador1))
    defuzzificacaoElevador1 = defuzzificacao(normalizacao, saidaFinalElevador1)
    
    resultadoFuzzi = defuzzificacaoElevador1

End Function
