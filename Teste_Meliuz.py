# para correta compilação desse código deve-se ter instalado: python, biblioteca pandas e o módulo openpyxl
import pandas as pd
from Graficos import Graficos_func
from Metricas_categorias import Metricas
from Metricas_gerais import Gerais

# Função para obter o mês de uma linha


def obter_mes(linha):
    return linha['Mês']

# Função para carregar a tabela a partir de um arquivo


def carregar_tabela(nome_arquivo):
    return pd.read_excel(nome_arquivo)

# Função para agrupar as linhas da tabela por categoria


def agrupar_por_categoria(tabela):
    linhas_por_categoria = {}
    for índice, linha in tabela.iterrows():
        categoria = linha['Categoria']
        if categoria in linhas_por_categoria:
            linhas_por_categoria[categoria].append(linha)
        else:
            linhas_por_categoria[categoria] = [linha]
    return ordenar_por_mes(linhas_por_categoria)

# Função para ordenar as linhas dentro de cada categoria pelo mês


def ordenar_por_mes(linhas_por_categoria):
    for categoria, linhas in linhas_por_categoria.items():
        linhas.sort(key=obter_mes)
    return linhas_por_categoria

# Função principal


def main():
    nome_arquivo = "Planilha_De_Dados.xlsx"
    tabela_excel = carregar_tabela(nome_arquivo)
    linhas_por_categoria = agrupar_por_categoria(tabela_excel)

    # gera os gráficos com as métricas para cada categoria presente na planilha
    for categoria_selecionada in linhas_por_categoria:
        tabela = Metricas.Tabela_qnt_vendas_e_cancelamentos(
            linhas_por_categoria, categoria_selecionada)

        tabela2 = Metricas.Tabela_numero_de_reclamacoes(
            linhas_por_categoria, categoria_selecionada)

        Graficos_func.Grafico_de_Vendas_Cancelamentos(
            tabela, categoria_selecionada)
        Graficos_func.criar_grafico_tendencia_vendas(
            tabela, categoria_selecionada)
        Metricas.Calculo_Taxas_e_Nova_Tabela(nome_arquivo)
        Graficos_func.Grafico_reclamacoes(tabela2, categoria_selecionada)

    # gera uma nova planilha de excel com as taxas de confirmação/validação por parceiro/mes/categoria
    Metricas.Calculo_Taxas_e_Nova_Tabela(nome_arquivo)

   # gera os gráficos com as métricas gerais da planilha
    Gerais.categorias_mais_vendidas_e_grafico(linhas_por_categoria)
    Gerais.categorias_maior_taxa_reclamacao_e_grafico(linhas_por_categoria)
    Gerais.grafico_meses_mais_reclamados(linhas_por_categoria)
    Gerais.parceiro_mais_reclamado_por_mes(linhas_por_categoria)
    
    # Relação entre taxas e reclamações
    Gerais.tabela_taxas_total_Reclamacoes(nome_arquivo)
    nome_tabela_correlacoes = 'Nova_Tabela_Taxas_e_Reclamacoes.xlsx'
    tabela_correlacoes = carregar_tabela(nome_tabela_correlacoes)
    Gerais.calcular_relacao_e_plotar_grafico(tabela_correlacoes)


if __name__ == "__main__":
    main()
