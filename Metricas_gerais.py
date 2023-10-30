# classe responsável por armazenar as funções responsáveis pelos gráficos gerais
import pandas as pd
import matplotlib.pyplot as plt
import statsmodels.api as sm
import seaborn as sns


class Gerais:

    # Função para calcular as categorias que mais venderam e gerar o gráfico
    def categorias_mais_vendidas_e_grafico(linhas_por_categoria):
        vendas_por_categoria = {}

        for categoria, linhas in linhas_por_categoria.items():
            total_vendas = sum(linha['Nº de vendas confirmadas']
                               for linha in linhas)
            vendas_por_categoria[categoria] = total_vendas

        categorias_ordenadas = sorted(
            vendas_por_categoria, key=lambda cat: vendas_por_categoria[cat], reverse=True)

        # Inverter a ordem das categorias
        categorias_ordenadas = categorias_ordenadas[::-1]

        categorias = categorias_ordenadas
        quantidades_vendidas = [vendas_por_categoria[cat]
                                for cat in categorias]

        # Normalizar os valores para que sejam proporcionais ao tamanho da imagem
        max_quantidade = max(quantidades_vendidas)
        norm_quantidades = [q / max_quantidade for q in quantidades_vendidas]

        plt.figure(figsize=(8, 6), dpi=100)
        plt.barh(categorias, norm_quantidades, color='skyblue')

        # Adicionar rótulos no topo das barras
        for i, v in enumerate(norm_quantidades):
            plt.text(
                v + 0.01, i, f'{quantidades_vendidas[i]}', ha='left', va='center', fontsize=8)

        # Diminuir o tamanho da fonte dos rótulos
        plt.tick_params(axis='y', labelsize=8)

        plt.xlabel('Proporção da Quantidade Total Vendida')
        plt.title('Categorias que mais venderam')

        plt.grid(axis='x', linestyle='--', linewidth=0.5)

        plt.show()

    # Função para calcular as categorias com a maior taxa de reclamação (considerando cancelamento e pendentes) e gerar o gráfico
    def categorias_maior_taxa_reclamacao_e_grafico(linhas_por_categoria):

        taxas_reclamacao_por_categoria = {}

        for categoria, linhas in linhas_por_categoria.items():
            total_reclamacoes_canc = sum(
                linha['Nº de reclamações por compra cancelada'] for linha in linhas)
            total_reclamacoes_pend = sum(
                linha['Nº de reclamações por compra pendente'] for linha in linhas)
            total_reclamacoes = total_reclamacoes_canc + total_reclamacoes_pend
            total_vendas = sum(linha['Nº de vendas confirmadas']
                               for linha in linhas)

            # Calcular a taxa de reclamação (reclamações / vendas)
            taxa_reclamacao = total_reclamacoes / total_vendas if total_vendas > 0 else 0
            taxas_reclamacao_por_categoria[categoria] = taxa_reclamacao

        # Ordenar as categorias com base na taxa de reclamação (do maior para o menor)
        categorias_ordenadas = sorted(taxas_reclamacao_por_categoria,
                                      key=lambda cat: taxas_reclamacao_por_categoria[cat], reverse=True)

        # Inverter a ordem das categorias
        categorias_ordenadas = categorias_ordenadas[::-1]

        categorias = categorias_ordenadas
        taxas_reclamacao = [taxas_reclamacao_por_categoria[cat]
                            for cat in categorias]

        plt.figure(figsize=(8, 6), dpi=100)
        plt.barh(categorias, taxas_reclamacao, color='lightcoral')

       # Adicionar rótulos no topo das barras
        for i, v in enumerate(taxas_reclamacao):
            # Adicionar as taxas de reclamação ao final de cada barra
            plt.text(v, i, f'{v:.2%}', ha='right',
                     va='center', fontsize=10, color='black')

        # Diminuir o tamanho da fonte dos rótulos
        plt.tick_params(axis='y', labelsize=8)

        plt.xlabel('Taxa de Reclamação')
        plt.title('Categorias com Maior Taxa de Reclamação')
        plt.grid(axis='x', linestyle='--', linewidth=0.5)

        plt.show()

    # função responsável por calcular e gerar o grafico para analise dos meses mais reclamados
    def grafico_meses_mais_reclamados(linhas_por_categoria):
        meses_reclamados = {}

        for categoria, linhas in linhas_por_categoria.items():
            for linha in linhas:
                mes = linha['Mês']
                reclamacoes_canceladas = linha['Nº de reclamações por compra cancelada']
                reclamacoes_pendentes = linha['Nº de reclamações por compra pendente']
                total_reclamacoes = reclamacoes_canceladas + reclamacoes_pendentes

                if mes in meses_reclamados:
                    meses_reclamados[mes] += total_reclamacoes
                else:
                    meses_reclamados[mes] = total_reclamacoes

        meses_ordenados = sorted(meses_reclamados.keys())
        reclamacoes_mensais = [meses_reclamados[mes]
                               for mes in meses_ordenados]

        plt.figure(figsize=(10, 6))
        plt.plot(meses_ordenados, reclamacoes_mensais,
                 marker='o', linestyle='-', color='b')
        plt.xlabel('Mês')
        plt.ylabel('Total de Reclamações')
        plt.title('Meses com Mais Reclamações')
        plt.xticks(rotation=45)
        plt.grid(True)

        plt.show()

    # função responsável encontrar o parceiro com maior nª de reclamações(canceladas+pendentes) por mes
    def parceiro_mais_reclamado_por_mes(linhas_por_categoria):
        parceiros_mais_reclamados = {}

        for categoria, linhas in linhas_por_categoria.items():
            for linha in linhas:
                mes = linha['Mês']
                parceiro = linha['Parceiro']
                reclamacoes_canceladas = linha['Nº de reclamações por compra cancelada']
                reclamacoes_pendentes = linha['Nº de reclamações por compra pendente']
                total_reclamacoes = reclamacoes_canceladas + reclamacoes_pendentes

                if mes not in parceiros_mais_reclamados:
                    parceiros_mais_reclamados[mes] = {
                        'parceiro': parceiro,
                        'total_reclamacoes': total_reclamacoes,
                        'categoria': categoria
                    }
                else:
                    if total_reclamacoes > parceiros_mais_reclamados[mes]['total_reclamacoes']:
                        parceiros_mais_reclamados[mes] = {
                            'parceiro': parceiro,
                            'total_reclamacoes': total_reclamacoes,
                            'categoria': categoria
                        }

        # Extraindo informações para o gráfico
        meses = list(parceiros_mais_reclamados.keys())
        parceiros = [parceiros_mais_reclamados[mes]['parceiro']
                     for mes in meses]
        total_reclamacoes = [parceiros_mais_reclamados[mes]
                             ['total_reclamacoes'] for mes in meses]

        # Criando o gráfico de barras
        plt.figure(figsize=(12, 8))
        x = range(len(meses))
        plt.bar(x, total_reclamacoes, align='center')
        plt.xlabel('Mês')
        plt.ylabel('Total de Reclamações')
        plt.title('Parceiro Mais Reclamado por Mês')
        plt.xticks(x, meses, rotation=45)
        plt.grid(True)

        # Anotando os parceiros no gráfico
        for i, total in enumerate(total_reclamacoes):
            plt.text(
                i, total + 10, f'Parceiro: {parceiros[i]}\nCategoria: {parceiros_mais_reclamados[meses[i]]["categoria"]}\nTotal: {total}', ha='center', va='bottom', fontsize=8)

        plt.show()

    # função responsável por criar uma nova tabela que armazena tanto as taxas quanto o total de reclamações,
    #  essa tabela será util para a análise da relação entre as taxas e as reclamações
    def tabela_taxas_total_Reclamacoes(nome_arquivo):

        # Carrega a tabela original
        tabela_excel = pd.read_excel(nome_arquivo)

    # Cálculo da Taxa de Confirmação
        tabela_excel['Taxa de Confirmação'] = (
            tabela_excel['Nº de vendas confirmadas'] / tabela_excel['Nº total de vendas'])

    # Cálculo da Taxa de Validação
        tabela_excel['Taxa de Validação'] = (
            (tabela_excel['Nº de vendas confirmadas'] + tabela_excel['Nº de vendas canceladas']) /
            tabela_excel['Nº total de vendas'])

    # Cálculo do Número Total de Reclamações
        tabela_excel['Nº de reclamações totais'] = (
            tabela_excel['Nº de vendas canceladas'] + tabela_excel['Nº de vendas pendentes'])

    # Seleciona as colunas desejadas
        nova_tabela = tabela_excel[[
            'Parceiro', 'Mês', 'Categoria', 'Taxa de Validação', 'Taxa de Confirmação', 'Nº de reclamações totais']]

    # Salva a nova tabela em um arquivo Excel
        nova_tabela.to_excel(
            'Nova_Tabela_Taxas_e_Reclamacoes.xlsx', index=False)

    # função responsável por gerar o gráfico de densidade que relaciona as taxas e reclamações
    def calcular_relacao_e_plotar_grafico(dicionario):
        # Criar um DataFrame a partir do dicionário
        df = pd.DataFrame(dicionario)

    # Calcular a relação entre as taxas e o número de reclamações totais
        df['Relação Taxa/Reclamações'] = df['Taxa de Validação'] / \
            df['Nº de reclamações totais']

    # Criar um gráfico de densidade
        plt.figure(figsize=(10, 6))
        sns.kdeplot(data=df, x='Taxa de Validação',
                    y='Nº de reclamações totais', fill=True, cmap='viridis')
        plt.xlabel('Taxa de Validação')
        plt.ylabel('Nº de reclamações totais')
        plt.title(
            'Distribuição da relação entre Taxa de Validação e Número de Reclamações Totais')
        plt.show()
