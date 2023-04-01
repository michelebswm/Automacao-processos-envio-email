# # Automação de Indicadores
#
# Criando um Projeto Completo de forma a gerar indicadores e enviar OnePage via e-mails para gerentes das 25 lojas.
#
# Um One Page é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os
# principais indicadores de cada loja e permitir em 1 página (daí o nome OnePage) tanto a comparação entre diferentes
# lojas, quanto quais indicadores aquela loja conseguiu cumprir naquele dia ou não.
#

# Arquivos e Informações Importantes
#
# - Arquivo Emails.xlsx com o nome, a loja e o e-mail de cada gerente.
#
# - Arquivo Vendas.xlsx com as vendas de todas as lojas. Obs: Cada gerente só deve receber o OnePage e um arquivo em
# excel em anexo com as vendas da sua loja. As informações de outra loja não devem ser enviados ao gerente que não é
# daquela loja.
#
# - Arquivo Lojas.csv com o nome de cada Loja
#
# - Ao final, sua rotina deve enviar ainda um e-mail para a diretoria (informações também estão no arquivo Emails.
# xlsx) com 2 rankings das lojas em anexo, 1 ranking do dia e outro ranking anual. Além disso, no corpo do e-mail,
# deve ressaltar qual foi a melhor e a pior loja do dia e também a melhor e pior loja do ano. O ranking de uma loja
# é dado pelo faturamento da loja.
#
# - As planilhas de cada loja devem ser salvas dentro da pasta da loja com a data da planilha, a fim de criar um
# histórico de backup
#
# Indicadores do OnePage
#
# - Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
# - Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4
# - Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500
#
# Obs: Cada indicador deve ser calculado no dia e no ano. O indicador do dia deve ser o do último dia disponível na
# planilha de Vendas (a data mais recente)

# ### Importando as Bibliotecas

import pandas as pd
import os
from pathlib import Path
import win32com.client as win32  # pip install pywin32  Biblioteca para enviar email Outlook

caminho = os.getcwd()
print(caminho)

# Importação da base de dados
vendas_df = pd.read_excel(rf'{caminho}\Bases de Dados\Vendas.xlsx')
lojas_df = pd.read_csv(rf'{caminho}\Bases de Dados\Lojas.csv', sep=';', encoding='latin1')
emails_df = pd.read_excel(rf'{caminho}\Bases de Dados\Emails.xlsx')
vendas_df = vendas_df.merge(lojas_df, on='ID Loja')
vendas_df = vendas_df.merge(emails_df, on='Loja')
print(vendas_df)

print(vendas_df.info())
print('-' * 60)
print(vendas_df.iloc[0])  # Analisando a primeiroa linha para entender os dados se estão corretos.

# ### Criando pastas e separando os arquivos por Loja em csv

ultimo_dia_vendas = str(vendas_df['Data'].max()).split(' ')[0].replace('-', '_')[
                    5:]  # Pega o ultimo dia de venda para incluir no nome do arquivo.
pasta = Path(caminho + r'\Backup Arquivos Lojas')

for loja in vendas_df['Loja'].unique():  # Pego o nome das lojas de forma única
    nova_pasta = pasta / loja  # Nome da Pasta incluindo o nome da Loja
    filtro = vendas_df.loc[vendas_df['Loja'] == loja,]  # Filtro por Loja
    if not nova_pasta.is_dir():  # Se a pasta não existe, cria a pasta
        nova_pasta.mkdir()
    if not (nova_pasta / Path(
            f'{ultimo_dia_vendas}_{loja}.csv')).exists():  # Se o arquivo não existe ele acessa o caminho da nova pasta e cria o arquivo csv com o filtro do Dataframe
        filtro.to_csv(rf'{nova_pasta}\{ultimo_dia_vendas}_{loja}.csv', sep=';', index=False)
    else:
        print('Existe')

    # ### Realizando análise dos indicadores

# Indicadores
meta_produtos_vendidos_diario = 4
meta_produtos_vendidos_anual = 120
meta_anual = 1650000
meta_diaria = 1000
meta_ticket_medio_diario = 500
meta_ticket_medio_anual = 500

# ### Pegando o ultimo dia do ano

ultimo_dia_vendas2 = str(vendas_df['Data'].max()).split(' ')[0]
print(ultimo_dia_vendas2)

# ### Criando dicionário de Lojas

dicionario_lojas = {}

for loja in lojas_df['Loja']:
    dicionario_lojas[loja] = vendas_df.loc[vendas_df['Loja'] == loja,]

print(dicionario_lojas['Iguatemi Esplanada'])

# ### Lógica

for loja in lojas_df['Loja']:
    # Filtro por Loja específica
    vendas_por_loja = vendas_df.loc[vendas_df['Loja'] == loja,]
    vendas_dia_especifico = vendas_por_loja.loc[vendas_por_loja['Data'] == ultimo_dia_vendas2]
    print(vendas_por_loja)
    print(vendas_dia_especifico)

    # - Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4
    diversidade_diaria = len(vendas_dia_especifico['Produto'].unique())
    diverdidade_anual = len(vendas_por_loja['Produto'].unique())
    print(diversidade_diaria, diverdidade_anual)

    # - Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
    faturamento_diario = vendas_dia_especifico['Valor Final'].sum()
    faturamento_anual = vendas_por_loja['Valor Final'].sum()
    print(faturamento_diario, faturamento_anual)

    # # - Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500
    total_venda_dia = vendas_dia_especifico.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_diario = total_venda_dia['Valor Final'].mean()

    total_venda_ano = vendas_por_loja.groupby('Código Venda').sum(numeric_only=True)
    tiket_medio_anual = total_venda_ano['Valor Final'].mean()
    print('ticket_medio_diario', ticket_medio_diario)
    print('tiket_medio_anual', tiket_medio_anual)

    # VALIDAÇÃO INDICADORES (META)
    diversidade_qtde_cor_dia = 'green' if diversidade_diaria >= meta_produtos_vendidos_diario else 'red'
    diversidade_qtde_cor_ano = 'green' if diverdidade_anual >= meta_produtos_vendidos_anual else 'red'

    faturamento_cor_dia = 'green' if faturamento_diario >= meta_diaria else 'red'
    faturamento_cor_ano = 'green' if faturamento_anual >= meta_anual else 'red'

    ticket_medio_cor_dia = 'green' if ticket_medio_diario >= meta_ticket_medio_diario else 'red'
    ticket_medio_cor_ano = 'green' if tiket_medio_anual >= meta_ticket_medio_anual else 'red'

    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = vendas_por_loja['E-mail'].unique()[0]  # E-mail do destinatário
    email.Subject = f'OnePage do dia {ultimo_dia_vendas}, Loja: {loja}'  # Assunto do E-mail
    nome_gerente = vendas_por_loja['Gerente'].unique()[0]
    # email.Body = "Mensagem do corpo do e-mail"
    email.HTMLBody = f'''
    <p> Bom dia, {nome_gerente}</p>

    <p>O resultado do dia {ultimo_dia_vendas} da Loja {loja} foi:</p>

    <table>
        <thead>
            <tr> 
                <td><b>Indicador</b></td>
                <td><b>Valor Dia</b></td>
                <td><b>Meta Dia</b></td>
                <td><b>Cenário Dia</b></td>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td style="text-align: center">Faturamento</td>
                <td style="text-align: center">{faturamento_diario:,.2f}</td>
                <td style="text-align: center">{meta_diaria:,.2f}</td>
                <td style="text-align: center"><font color="{faturamento_cor_dia}">◙</font></td>
            </tr>
            <tr>
                <td style="text-align: center">Diversidade</td>
                <td style="text-align: center">{diversidade_diaria}</td>
                <td style="text-align: center">{meta_produtos_vendidos_diario}</td>
                <td style="text-align: center"><font color="{diversidade_qtde_cor_dia}">◙</font></td>
            </tr>
            <tr>
                <td style="text-align: center">Ticket Médio</td>
                <td style="text-align: center">{ticket_medio_diario:,.2f}</td>
                <td style="text-align: center">{meta_ticket_medio_diario:,.2f}</td>
                <td style="text-align: center"><font color="{ticket_medio_cor_dia}">◙</font></td>
            </tr>
        </tbody>
    </table>
    <table>
        <thead>
            <tr> 
                <td><b>Indicador</b></td>
                <td><b>Valor Ano</b></td>
                <td><b>Meta Ano</b></td>
                <td><b>Cenário Ano</b></td>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td style="text-align: center">Faturamento</td>
                <td style="text-align: center">{faturamento_anual:,.2f}</td>
                <td style="text-align: center">{meta_anual:,.2f}</td>
                <td style="text-align: center"><font color="{faturamento_cor_ano}">◙</font></td>
            </tr>
            <tr>
                <td style="text-align: center">Diversidade</td>
                <td style="text-align: center">{diverdidade_anual}</td>
                <td style="text-align: center">{meta_produtos_vendidos_anual}</td>
                <td style="text-align: center"><font color="{diversidade_qtde_cor_ano}">◙</font></td>
            </tr>
            <tr>
                <td style="text-align: center">Ticket Médio</td>
                <td style="text-align: center">{tiket_medio_anual:,.2f}</td>
                <td style="text-align: center">{meta_ticket_medio_anual:,.2f}</td>
                <td style="text-align: center"><font color="{ticket_medio_cor_ano}">◙</font></td>
            </tr>
        </tbody>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes</p>
    <p>Qualquer dúvida, estou a disposição.</p>


    <p>Att.,</p>
    <p>Fulano</p>


    '''
    attachment = pasta / loja / f'{ultimo_dia_vendas}_{loja}.csv'
    print(attachment)

    email.Attachments.Add(str(attachment))  # Para adicionar tem que transformar em string

    email.Send()  # Enviar o e-mail
    print('Email enviado para {} referente a loja {}'.format(nome_gerente, loja))

# ### Análise dia a dia criando coluna Atingiu Meta como Sim ou Não

# Analise Anual dia por dia
for loja in lojas_df['Loja']:
    # Filtro por Loja específica
    vendas_por_loja = vendas_df.loc[vendas_df['Loja'] == loja,]
    print(vendas_por_loja)

    # - Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4
    vendas_por_data = vendas_por_loja.groupby(['Data'])['Produto'].unique().reset_index()
    for i in range(0, len(vendas_por_data)):
        vendas_por_data.loc[i, 'Tot Prod Diferentes'] = len(vendas_por_data['Produto'].iloc[i])
        if vendas_por_data['Tot Prod Diferentes'][i] >= meta_produtos_vendidos_diario:
            vendas_por_data.loc[i, 'Atingiu Meta'] = 'Sim'
        else:
            vendas_por_data.loc[i, 'Atingiu Meta'] = 'Nao'

    print('Produtos diferentes vendidos no ano: ', len(vendas_por_loja['Produto'].unique()))
    print(vendas_por_data)

    # - Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
    # Faturamento por Dia
    faturamento_diario = vendas_por_loja.groupby('Data')['Valor Final'].sum().reset_index()

    for i, valor in enumerate(faturamento_diario['Valor Final']):
        if valor >= meta_diaria:
            faturamento_diario.loc[i, 'Atingiu Meta'] = 'Sim'  # Uso o .loc na linha i, coluna 'Atingiu Meta'
        else:
            faturamento_diario.loc[i, 'Atingiu Meta'] = 'Não'

    print(faturamento_diario)

    # Faturamento Anual
    faturamento_anual = vendas_por_loja['Valor Final'].sum()
    print("Faturamento Anual: R$ {:,.2f}".format(faturamento_anual))

    # - Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500
    ticket_medio_diario = vendas_por_loja.groupby('Data')['Valor Final'].mean().reset_index()
    for i in range(0, len(ticket_medio_diario['Valor Final'])):
        if ticket_medio_diario['Valor Final'][i] > meta_ticket_medio_diario:
            ticket_medio_diario.loc[i, 'Atingiu Meta'] = 'Sim'
        else:
            ticket_medio_diario.loc[i, 'Atingiu Meta'] = 'Nao'
    print(ticket_medio_diario)

    tiket_medio_anual = vendas_por_loja['Valor Final'].mean()
    print('Ticket Medio Anual: ', tiket_medio_anual)

# ### Ranking de Faturamento

faturamento_por_loja = vendas_df.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True).reset_index()
faturamento_anual_ordenado = faturamento_por_loja.sort_values(by='Valor Final',
                                                              ascending=False)  # Ordenando de forma decrescente
print(faturamento_anual_ordenado)

# Exportar excel Ranking_Anual Faturamento
faturamento_anual_ordenado.to_excel(rf'{pasta}\{ultimo_dia_vendas}_Ranking_Anual.xlsx', index=False)

faturamento_loja_dia = vendas_df.loc[vendas_df['Data'] == ultimo_dia_vendas2,]
fat_dia = faturamento_loja_dia.groupby('Loja')[['Loja', 'Valor Final']].sum(
    numeric_only=True).reset_index()  # Ordenando de forma decrescente
faturamento_diario_ordenado = fat_dia.sort_values(by='Valor Final', ascending=False)
print(faturamento_diario_ordenado)

# Exportar excel Ranking_Dia Faturamento
faturamento_diario_ordenado.to_excel(rf'{pasta}\{ultimo_dia_vendas}_Ranking_Dia.xlsx', index=False)

email_diretoria = [email for email in emails_df['E-mail'] if 'diretoria' in email]
print(email_diretoria[0])

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')
# criar um email
email = outlook.CreateItem(0)

# configurar as informações do seu e-mail
email.To = email_diretoria[0]  # Email Destino
email.Subject = "Ranking de Faturamento Diario e Anual Lojas"
email.Body = f'''
A Loja que obteve o maior faturamento no dia {ultimo_dia_vendas2} foi {faturamento_diario_ordenado['Loja'].values[0]} com um faturamento anual de R$ {faturamento_diario_ordenado['Valor Final'].values[0]:,.2f}
A Loja que obteve o pior faturamento no dia {ultimo_dia_vendas2} foi {faturamento_diario_ordenado['Loja'].values[-1]} com um faturamento anual de R$ {faturamento_diario_ordenado['Valor Final'].values[-1]:,.2f}

A Loja que obteve o maior faturamento anual foi {faturamento_anual_ordenado['Loja'].values[0]} com um faturamento anual de R$ {faturamento_anual_ordenado['Valor Final'].values[0]:,.2f}
A loja que obteve o pior faturamento anual foi  {faturamento_anual_ordenado['Loja'].values[-1]} com um faturamento anual de R$ {faturamento_anual_ordenado['Valor Final'].values[-1]:,.2f}
'''

anexo1 = pasta / f'{ultimo_dia_vendas}_Ranking_Dia.xlsx'
anexo2 = pasta / f'{ultimo_dia_vendas}_Ranking_Anual.xlsx'
email.Attachments.Add(str(anexo1))
email.Attachments.Add(str(anexo2))

email.Send()
print('E-mail da Diretoria enviado')



