import numpy as np
import matplotlib.pyplot as plt
import skfuzzy.membership as mf

normalizacao = np.arange(0, 1, 10**(-2))

tempoEsperaBaixo = mf.trapmf(normalizacao, [0, 0, 0.3, 0.5])
tempoEsperaMedio = mf.trimf(normalizacao, [0.25, 0.55, 0.85])
tempoEsperaAlto = mf.trimf(normalizacao, [0.6, 1, 1])

distanciaBaixo = mf.trapmf(normalizacao, [0, 0, 0.3, 0.5])
distanciaMedio = mf.trimf(normalizacao, [0.25, 0.55, 0.85])
distanciaAlto = mf.trimf(normalizacao, [0.6, 1, 1])

capacidadeBaixo = mf.trimf(normalizacao,[0, 0, 0.4])
capacidadeMedio = mf.trimf(normalizacao,[0.2, 0.45, 0.8])
capacidadeAlto = mf.trimf(normalizacao,[0.7, 1, 1])

prioridadeBaixo = mf.trimf(normalizacao,[0, 0, 0.4])
prioridadeMedio = mf.trimf(normalizacao,[0.1, 0.5, 0.9])
prioridadeAlto = mf.trimf(normalizacao,[0.6, 1, 1])

plt.plot(normalizacao, tempoEsperaBaixo, 'r', linewidth = 2, label = 'Baixo')
plt.plot(normalizacao, tempoEsperaMedio, 'g', linewidth = 2, label = 'Médio')
plt.plot(normalizacao, tempoEsperaAlto, 'b', linewidth = 2, label = 'Alto')
plt.xlabel("Tempo de Espera Normalizado")
plt.ylabel("Grau de pertinência")
plt.legend()

plt.tight_layout()
plt.show()


plt.plot(normalizacao, distanciaBaixo, 'r', linewidth = 2, label = 'Baixa')
plt.plot(normalizacao, distanciaMedio, 'g', linewidth = 2, label = 'Média')
plt.plot(normalizacao, distanciaAlto, 'b', linewidth = 2, label = 'Alta')
plt.xlabel("Distância Normalizada")
plt.ylabel("Grau de pertinência")
plt.legend()

plt.tight_layout()
plt.show()

plt.plot(normalizacao, capacidadeBaixo, 'r', linewidth = 2, label = 'Baixa')
plt.plot(normalizacao, capacidadeMedio, 'g', linewidth = 2, label = 'Média')
plt.plot(normalizacao, capacidadeAlto, 'b', linewidth = 2, label = 'Alta')
plt.xlabel("Capacidade Normalizada")
plt.ylabel("Grau de pertinência")
plt.legend()

plt.tight_layout()
plt.show()

plt.plot(normalizacao, prioridadeBaixo, 'r', linewidth = 2, label = 'Baixa')
plt.plot(normalizacao, prioridadeMedio, 'g', linewidth = 2, label = 'Média')
plt.plot(normalizacao, prioridadeAlto, 'b', linewidth = 2, label = 'Alta')
plt.xlabel("Prioridade Normalizada")
plt.ylabel("Grau de pertinência")
plt.legend()

plt.tight_layout()
plt.show()
