# classe responsável por armazenar as funções responsáveis pelos gráficos das categorias da planilha
import pandas as pd
import matplotlib.pyplot as plt
import statsmodels.api as sm


class Graficos_func:

    # função responsável pelo gráfico de vendas e cancelamentos de cada mês de cada catégoria
    def Grafico_de_Vendas_Cancelamentos(tabela, categoria):
        plt.figure(figsize=(10, 6))
        largura_barra = 0.40
        indice = tabela.index
        azul_claro = '#3498db'

        plt.bar(indice, tabela['Total de Vendas'],
                largura_barra, label='Vendas', color=azul_claro)
        plt.bar(indice + largura_barra, tabela['Total de Cancelamentos'],
                largura_barra, label='Cancelamentos', color='red')

        plt.xlabel('Mês')
        plt.ylabel('Quantidade')
        plt.title(
            f'Quantidade Total de Vendas e Cancelamentos por Mês - {categoria}')

        for i, v in enumerate(tabela['Total de Vendas']):
            plt.text(i, v + 10, str(v), ha='center', va='bottom', fontsize=8)

        for i, c in enumerate(tabela['Total de Cancelamentos']):
            plt.text(i + largura_barra, c + 10, str(c),
                     ha='center', va='bottom', fontsize=8)

        plt.xticks(indice + largura_barra / 2, tabela['Mês'], fontsize=8)
        plt.legend()
        plt.tight_layout()
        plt.show()

    # funçã responsável por criar o gráfico das reclamações em cada mês de cada categoria
    def Grafico_reclamacoes(tabela, categoria):
        plt.figure(figsize=(10, 6))

        largura_barra = 0.40  # Largura das barras
        indice = tabela.index  # Índice das barras
        azul_claro = '#3498db'

        plt.bar(indice, tabela['Total de reclamações por compra pendente'],
                largura_barra, label='Reclamações compra pendente', color=azul_claro)
        plt.bar(indice + largura_barra, tabela['Total de reclamações por compra cancelada'],
                largura_barra, label='Reclamações compra cancelada', color='red')

        plt.xlabel('Mês')
        plt.ylabel('Quantidade')
        plt.title(f'Quantidade Total de Reclamações por Mês - {categoria}')

        for i, v1, v2 in zip(indice, tabela['Total de reclamações por compra pendente'], tabela['Total de reclamações por compra cancelada']):
            plt.text(i, v1 + 10, str(v1), ha='center', va='bottom', fontsize=8)
            plt.text(i + largura_barra, v2 + 10, str(v2),
                     ha='center', va='bottom', fontsize=8)

        # Rotação dos rótulos dos meses em 45 graus
        plt.xticks(indice + largura_barra / 2, tabela['Mês'], fontsize=8)
        plt.legend()
        plt.tight_layout()
        plt.show()

    # função responsável por criar os gráficos de taxa de crescimento e de tendencia de vendas no período de análise para cada categoria
    def criar_grafico_tendencia_vendas(tabela, categoria):
        # Crie uma nova coluna para ordenação (assumindo que os valores são inteiros)
        tabela['Mes_Ordenacao'] = tabela['Mês'].astype(int)

        # Classificar os dados pelo mês de ordenação
        tabela = tabela.sort_values(by='Mes_Ordenacao')

        # Calcular a taxa de crescimento das vendas
        tabela['Taxa de Crescimento'] = tabela['Total de Vendas'].pct_change() * \
            100

        # gráfico da tendencia de vendas
        plt.figure(figsize=(10, 6))
        plt.plot(tabela['Mês'], tabela['Total de Vendas'],
                 label='Vendas', marker='o')
        plt.xlabel('Mês')
        plt.ylabel('Total de Vendas')
        plt.title(f'Tendência de Vendas ao longo do Tempo - {categoria}')
        plt.xticks(rotation=45)
        plt.grid(True)
        plt.legend()

        # Gráfico para a taxa de crescimento
        plt.figure(figsize=(10, 6))
        plt.plot(tabela['Mês'], tabela['Taxa de Crescimento'],
                 label='Taxa de Crescimento', marker='o', color='red')
        plt.xlabel('Mês')
        plt.ylabel('Taxa de Crescimento (%)')
        plt.title(f'Taxa de Crescimento de Vendas - {categoria}')
        plt.xticks(rotation=45)
        plt.grid(True)
        plt.legend()
        plt.show()
