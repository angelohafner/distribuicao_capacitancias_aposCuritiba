import numpy as np
import pandas as pd
import random
""""
Funcões que uso no arquivo principal
"""


def criar_matrizes(num_capacitancias, capacitancias, num_serie_cap, capacitancias_por_matriz, n_lin, n_col):
    matrizes = []
    numeros = []
    for i in range(0, num_capacitancias, capacitancias_por_matriz):
        matriz = capacitancias[i:i + capacitancias_por_matriz].reshape(n_lin, n_col)
        numero = num_serie_cap[i:i + capacitancias_por_matriz].reshape(n_lin, n_col)
        matrizes.append(matriz)
        numeros.append(numero)

    return matrizes, numeros



def calcular_capacitancia_equivalente(matriz):
    soma_paralelo = np.sum(matriz, axis=0)
    inverso_serie = np.sum(1.0 / soma_paralelo)
    return 1.0 / inverso_serie


def diferenca_total_capacitancias_equivalentes(capacitancias):
    return np.ptp(capacitancias)


def trocar_capacitancias(matrizes, numeros, n_lin, n_col):
    """"
    idx1 e idx2 ==> localização de uma matriz n_lin x n_col
    pos1 e pos3 ==> tupla que representa a posição do elemento na matriz
    Obs: para ter acesso ao elemento da matriz precisamos de:
             - idx para localizar a matriz
             - pos para localizar o elemento
    """
    idx1, idx2 = random.sample(range(len(matrizes)), 2)
    pos1 = (random.randrange(n_lin), random.randrange(n_col))
    pos2 = (random.randrange(n_lin), random.randrange(n_col))
    # primeiro para as capacitâncias, depois para os números de série
    temp = matrizes[idx1][pos1]
    matrizes[idx1][pos1] = matrizes[idx2][pos2]
    matrizes[idx2][pos2] = temp
    temp = numeros[idx1][pos1]
    numeros[idx1][pos1] = numeros[idx2][pos2]
    numeros[idx2][pos2] = temp


def otimizar_capacitancias(n_iteracoes,
                            matrizes, numeros, n_lin, n_col,
                            calcular_capacitancia_equivalente):

    menor_diferenca = float('inf')

    # faz n_iteracoes trocando capacitâncias. Segura aquele cujo valor foi o melhor.
    for _ in range(n_iteracoes):
        # lembrar que a troca ocorre em uma lista que tem nível global,
        # por este motivo nenhuma variável precisa ser retornada
        trocar_capacitancias(matrizes, numeros, n_lin, n_col)
        capacitancias_equivalentes_atualizadas = []

        for matriz in matrizes:
            cap_equiv = calcular_capacitancia_equivalente(matriz)
            capacitancias_equivalentes_atualizadas.append(cap_equiv)

        diferenca_atual = diferenca_total_capacitancias_equivalentes(capacitancias_equivalentes_atualizadas)

        if diferenca_atual < menor_diferenca:
            menor_diferenca = diferenca_atual
            melhor_configuracao_value = [np.copy(matriz) for matriz in matrizes]
            melhor_configuracao_nrser = [np.copy(numero) for numero in numeros]

    return melhor_configuracao_value, melhor_configuracao_nrser


def exportar_matrizes_para_excel(melhor_configuracao_nrser, melhor_configuracao_value, arquivo_nrser, arquivo_value):
    num_matrizes = len(melhor_configuracao_nrser)

    with pd.ExcelWriter(arquivo_nrser) as writer1:
        with pd.ExcelWriter(arquivo_value) as writer2:
            for ind in range(num_matrizes):
                # Convertendo as matrizes em DataFrames
                df_nrser = pd.DataFrame(melhor_configuracao_nrser[ind])
                df_value = pd.DataFrame(melhor_configuracao_value[ind])

                # Exportando cada matriz para uma aba diferente no arquivo Excel
                df_nrser.to_excel(writer1, sheet_name=f'Ramo_{ind}')
                df_value.to_excel(writer2, sheet_name=f'Ramo_{ind}')
