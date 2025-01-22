#importar bibliotecas
import pandas as pd
import pathlib
import win32com.client as win32


#importar bases de dados

emails = pd.read_excel(r"D:\aulas de python\Currículo\Automação de Processos\Base de dados\Emails.xlsx")
lojas = pd.read_csv(r"D:\aulas de python\Currículo\Automação de Processos\Base de dados\Lojas.csv", encoding="latin1", sep=";")
vendas = pd.read_excel(r"D:\aulas de python\Currículo\Automação de Processos\Base de dados\Vendas.xlsx")

#incluir nome da loja em vendas
vendas = vendas.merge(lojas, on= 'ID Loja')

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

dia_indicador = vendas['Data'].max()
print('{}/{}'.format(dia_indicador.day, dia_indicador.month))

#identificar se a pasta já existe
caminho_backup = pathlib.Path(r'D:\aulas de python\Currículo\Automação de Processos\Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

    #salvar dentro da pasta
    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtndprodutos_dia = 4
meta_qtndprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

for loja in dicionario_lojas:
    
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja["Data"] == dia_indicador , :]

    # Faturamento 
    faturamento_ano = vendas_loja['Valor Final'].sum()

    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    # Diversidade de produtos
    qtnd_produtos_ano = len(vendas_loja['Produto'].unique())

    qtnd_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    # Ticket médio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    media_ano = valor_venda['Valor Final'].mean()

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    media_dia = valor_venda_dia['Valor Final'].mean()


    # Criar o email
    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja'] == loja ,'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja'] == loja ,'E-mail'].values[0]  #excluir o indice
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'

    # Automação da cor do indicador cenário
    cor_fat_dia = 'green' if faturamento_dia >= meta_faturamento_dia else 'red'
    cor_fat_ano = 'green' if faturamento_ano >= meta_faturamento_ano else 'red'
    cor_qtnd_dia = 'green' if qtnd_produtos_dia >= meta_qtndprodutos_dia else 'red'
    cor_qtnd_ano = 'green' if qtnd_produtos_ano >= meta_qtndprodutos_ano else 'red'
    cor_ticket_dia = 'green' if media_dia >= meta_ticketmedio_dia else 'red'
    cor_ticket_ano = 'green' if media_ano >= meta_ticketmedio_ano else 'red' 


    mail.HTMLBody = f'''
    <p>Bom dia, {nome}</p>

    <p> O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor dia</th>
        <th>Meta dia</th>
        <th>Cenário dia</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style='text-align: center'>R${faturamento_dia:.2f}</td>
        <td style='text-align: center'>R${meta_faturamento_dia:.2f}</td>
        <td style='text-align: center'><font color='{cor_fat_dia}'>◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de produtos</td>
        <td style='text-align: center'>{qtnd_produtos_dia}</td>
        <td style='text-align: center'>{meta_qtndprodutos_dia}</td>
        <td style='text-align: center'><font color='{cor_qtnd_dia}'>◙</font></td>
    </tr>
    <tr>
        <td style='text-align: center'>Ticket Médio</td>
        <td style='text-align: center'>R${media_dia:.2f}</td>
        <td style='text-align: center'>R${meta_ticketmedio_dia:.2f}</td>
        <td style='text-align: center'><font color='{cor_ticket_dia}'>◙</font></td>
    </tr>
    </table>
    <br>
    <table>
    <tr>
        <th>Indicador</th>
        <th>Valor ano</th>
        <th>Meta ano</th>
        <th>Cenário ano</th>
    </tr>
    <tr>
        <td>Faturamento</td>
        <td style='text-align: center'>R${faturamento_ano:.2f}</td>
        <td style='text-align: center'>R${meta_faturamento_ano:.2f}</td>
        <td style='text-align: center'><font color='{cor_fat_ano}'>◙</font></td>
    </tr>
    <tr>
        <td>Diversidade de produtos</td>
        <td style='text-align: center'>{qtnd_produtos_ano}</td>
        <td style='text-align: center'>{meta_qtndprodutos_ano}</td>
        <td style='text-align: center'><font color='{cor_qtnd_ano}'>◙</font></td>
    </tr>
    <tr>
        <td style='text-align: center'>Ticket Médio</td>
        <td style='text-align: center'>R${media_ano:.2f}</td>
        <td style='text-align: center'>R${meta_ticketmedio_ano:.2f}</td>
        <td style='text-align: center'><font color='{cor_ticket_ano}'>◙</font></td>
    </tr>
    </table>

    <p> Segue em anexo a planilha de todos os dados para mais detalhes.</p>
    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Vinícius.</p>
    '''

    # Anexos:
    attachment  = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    # Enviar o email
    mail.Send()
    print("email")

faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

#salvar dentro da pasta
nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


vendas_dia = vendas.loc[vendas["Data"] == dia_indicador , :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

#salvar dentro da pasta
nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

# Criar o email
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja'] == 'Diretoria' ,'E-mail'].values[0]  #excluir o indice
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0,0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1,0]:.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0,0]:.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1,0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Vinicius
'''

# Anexos:
attachment  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))

# Enviar o email
mail.Send()


