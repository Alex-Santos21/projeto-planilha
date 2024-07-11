import pandas as pd
import win32com.client as win32

# Tentar importar a base de dados
try:
    tabela_vendas = pd.read_excel('Vendas.xlsx')
except Exception as e:
    print(f"Erro ao ler o arquivo Excel: {e}")
    tabela_vendas = None

# Continuar somente se a tabela foi carregada com sucesso
if tabela_vendas is not None:
    # Verificar se as colunas necessárias estão presentes
    colunas_necessarias = {'ID Loja', 'Valor Final', 'Quantidade'}
    if not colunas_necessarias.issubset(tabela_vendas.columns):
        print(f"Erro: As colunas necessárias não estão presentes na tabela. Colunas encontradas: {tabela_vendas.columns}")
    else:
        # Visualizar a base de dados
        pd.set_option('display.max_columns', None)
        print(tabela_vendas)

        # Faturamento por loja
        faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
        print(faturamento)

        # Quantidade de produtos vendidos por loja
        quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
        print(quantidade)

        print('-' * 50)

        # Ticket médio por produto em cada loja
        ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame(name='Ticket Médio')
        print(ticket_medio)

        # Enviar um email com o relatório
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'pythonimpressionador@gmail.com'
            mail.Subject = 'Relatório de Vendas por Loja'
            mail.HTMLBody = f'''
            <p>Prezados,</p>

            <p>Segue o Relatório de Vendas por cada Loja.</p>

            <p>Faturamento:</p>
            {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

            <p>Quantidade Vendida:</p>
            {quantidade.to_html()}

            <p>Ticket Médio dos Produtos em cada Loja:</p>
            {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

            <p>Qualquer dúvida estou à disposição.</p>

            <p>Att.,</p>
            <p>Lira</p>
            '''
            mail.Send()
            print('Email Enviado')
        except Exception as e:
            print(f"Erro ao enviar o email: {e}")
