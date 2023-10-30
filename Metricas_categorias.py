# classe responsável por realizar as metricas em cada categoria da planilha de case
import pandas as pd


class Metricas:
    # Função para criar uma tabela(DataFrame) com quantidade total de vendas e cancelamentos por mês para cada categoria da planilha
    def Tabela_qnt_vendas_e_cancelamentos(linhas_por_categoria, categoria):

        meses = []  # Use uma lista para armazenar os meses
        total_vendas = []  # Inicialize uma lista vazia para total de vendas
        total_cancelamentos = []  # Inicialize uma lista vazia para total de cancelamentos

        for linha in linhas_por_categoria[categoria]:
            mes_ano = linha['Mês']
            vendas_confirmadas = linha['Nº de vendas confirmadas']
            vendas_canceladas = linha['Nº de vendas canceladas']

            # Extrai o mês (o último elemento após dividir o texto por "-")
            mes = mes_ano.split("-")[-1]

            if mes not in meses:
                meses.append(mes)
                total_vendas.append(vendas_confirmadas)
                total_cancelamentos.append(vendas_canceladas)
            else:
                idx = meses.index(mes)
                total_vendas[idx] += vendas_confirmadas
                total_cancelamentos[idx] += vendas_canceladas

        # Ordenar os meses
        meses = sorted(meses)

        tabela = pd.DataFrame({'Mês': meses, 'Total de Vendas': total_vendas,
                              'Total de Cancelamentos': total_cancelamentos})

        return tabela

    # Função para criar uma tabela(DataFrame) com quantidade total de reclamações por mês para cada categoria da planilha
    def Tabela_numero_de_reclamacoes(linhas_por_categoria, categoria):
        meses = []  # Use uma lista para armazenar os meses
        reclamacoes_cancelamento = []
        reclamacoes_pendentes = []

        for linha in linhas_por_categoria[categoria]:
            mes_ano = linha['Mês']
            reclamacoes_canc = linha['Nº de reclamações por compra cancelada']
            reclamacoes_pen = linha['Nº de reclamações por compra pendente']

            # Extrai o mês (o último elemento após dividir o texto por "-")
            mes = mes_ano.split("-")[-1]

            if mes not in meses:
                meses.append(mes)
                reclamacoes_pendentes.append(reclamacoes_pen)
                reclamacoes_cancelamento.append(reclamacoes_canc)
            else:
                idx = meses.index(mes)
                reclamacoes_pendentes[idx] += reclamacoes_pen
                reclamacoes_cancelamento[idx] += reclamacoes_canc

        # Ordenar os meses
        meses = sorted(meses)

        tabela_reclamacoes = pd.DataFrame({'Mês': meses, 'Total de reclamações por compra cancelada': reclamacoes_cancelamento,
                                          'Total de reclamações por compra pendente': reclamacoes_pendentes})

        return tabela_reclamacoes

    # função responsável por calcular as taxas de validação/confirmação por parceiro/mês/categoria e gerar uma nova tabela excel
    def Calculo_Taxas_e_Nova_Tabela(nome_arquivo):
        # Carrega a tabela original
        tabela_excel = pd.read_excel(nome_arquivo)

        # Cálculo da Taxa de Confirmação
        tabela_excel['Taxa de Confirmação'] = (
            tabela_excel['Nº de vendas confirmadas'] / tabela_excel['Nº total de vendas'])

        # Cálculo da Taxa de Validação
        tabela_excel['Taxa de Validação'] = ((tabela_excel['Nº de vendas confirmadas'] +
                                              tabela_excel['Nº de vendas canceladas']) / tabela_excel['Nº total de vendas'])

        # Seleciona as colunas desejadas
        nova_tabela = tabela_excel[[
            'Parceiro', 'Mês', 'Categoria', 'Taxa de Validação', 'Taxa de Confirmação']]

        # Salva a nova tabela em um arquivo Excel
        nova_tabela.to_excel('Nova_Tabela_Taxas.xlsx', index=False)
