Private Sub VBA_Block_2_Fire()

Dim s As SIMAN
Set s = ThisDocument.Model.SIMAN

Dim andarElevador1 As Double
Dim andarElevador2 As Double
Dim limiteCapacidadeElevador As Double
Dim numeroAndares As Double
Dim NumParadasElevador1 As Double
Dim NumParadasElevador2 As Double
Dim tempoVoo As Double
Dim tempoParada As Double
Dim tempoMaximo As Double
Dim distanciaElevador1Normalizada As Double
Dim distanciaElevador2Normalizada As Double
Dim tempoEsperaElevador1 As Double
Dim tempoEsperaElevador2 As Double
Dim cargaElevador1 As Double
Dim cargaElevador2 As Double
Dim cargaElevador1Normalizada As Double
Dim cargaElevador2Normalizada As Double

Dim andarOrigemValor As Integer
Dim elevadorSelecionadoValor As Integer
Dim andarDestinoValor As Integer
Dim isPessoaSubindoValor As Integer

limiteCapacidadeElevador = 6
numeroAndares = 6
tempoVoo = 2
tempoParada = 3
tempoMaximo = ((numeroAndares - 1) * tempoVoo) + (2 * numeroAndares * tempoParada)

andarOrigemValor = s.EntityAttribute(s.ActiveEntity, s.SymbolNumber("andarOrigem"))
andarDestinoValor = s.EntityAttribute(s.ActiveEntity, s.SymbolNumber("andarDestino"))
isPessoaSubindoValor = s.EntityAttribute(s.ActiveEntity, s.SymbolNumber("isPessoaSubindo"))

andarElevador1 = s.VariableArrayValue(s.SymbolNumber("estadoElevadores", 1, 4))
andarElevador2 = s.VariableArrayValue(s.SymbolNumber("estadoElevadores", 2, 4))

NumParadasElevador1 = 0
NumParadasElevador2 = 0
For i = 1 To numeroAndares
    If s.VariableArrayValue(s.SymbolNumber("paradasElevadoresSubindo", 1, i)) = 1 Then
        NumParadasElevador1 = NumParadasElevador1 + 1
    End If
    If s.VariableArrayValue(s.SymbolNumber("paradasElevadoresSubindo", 2, i)) = 1 Then
        NumParadasElevador2 = NumParadasElevador2 + 1
    End If
    If s.VariableArrayValue(s.SymbolNumber("paradasElevadoresDescendo", 1, i)) = 1 Then
        NumParadasElevador1 = NumParadasElevador1 + 1
    End If
    If s.VariableArrayValue(s.SymbolNumber("paradasElevadoresDescendo", 2, i)) = 1 Then
        NumParadasElevador2 = NumParadasElevador2 + 1
    End If
Next i

cargaElevador1 = s.VariableArrayValue(s.SymbolNumber("estadoElevadores", 1, 3))
cargaElevador2 = s.VariableArrayValue(s.SymbolNumber("estadoElevadores", 2, 3))

distanciaElevador1Normalizada = Abs(andarOrigemValor - andarElevador1) / (numeroAndares - 1)
distanciaElevador2Normalizada = Abs(andarOrigemValor - andarElevador2) / (numeroAndares - 1)

tempoEsperaElevador1 = ((Abs(andarOrigemValor - andarElevador1) * tempoVoo) + (2 * NumParadasElevador1 * tempoParada)) / tempoMaximo
tempoEsperaElevador2 = ((Abs(andarOrigemValor - andarElevador2) * tempoVoo) + (2 * NumParadasElevador2 * tempoEntra)) / tempoMaximo

cargaElevador1Normalizada = cargaElevador1 / limiteCapacidadeElevador
cargaElevador2Normalizada = cargaElevador2 / limiteCapacidadeElevador

elevadorSelecionadoValor = selecaoElevador(distanciaElevador1Normalizada, distanciaElevador2Normalizada, tempoEsperaElevador1, tempoEsperaElevador2, cargaElevador1Normalizada, cargaElevador2Normalizada)

s.EntityAttribute(s.ActiveEntity, s.SymbolNumber("elevadorSelecionado")) = elevadorSelecionadoValor

If isPessoaSubindoValor = 1 Then
    s.VariableArrayValue(s.SymbolNumber("paradasElevadoresSubindo", elevadorSelecionadoValor, andarOrigemValor)) = 1
    s.VariableArrayValue(s.SymbolNumber("paradasElevadoresSubindo", elevadorSelecionadoValor, andarDestinoValor)) = 1
Else
    s.VariableArrayValue(s.SymbolNumber("paradasElevadoresDescendo", elevadorSelecionadoValor, andarOrigemValor)) = 1
    s.VariableArrayValue(s.SymbolNumber("paradasElevadoresDescendo", elevadorSelecionadoValor, andarDestinoValor)) = 1
End If

End Sub
