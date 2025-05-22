#==============================================================================================
# Bibliotecas
#==============================================================================================
import os
import sys
import gspread
import funcoes
import pandas as pd
from dotenv import load_dotenv
from googleapiclient.errors import HttpError # capturar uma exceção HttpError
from oauth2client.service_account import ServiceAccountCredentials

#==============================================================================================
# Execução do programa
#==============================================================================================
if not funcoes.verificar_conexao():
   print(f"\nSem conexão com a internet.\n")
   sys.exit()

# Definir o escopo de acesso à API do Google Sheets
escopo = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

load_dotenv()

# Autenticar com as credenciais da conta de serviço
credenciais = ServiceAccountCredentials.from_json_keyfile_name(
        os.getenv('GOOGLE_CREDS_JSON_PATH'), escopo
)
if not credenciais:
   print(f"\nServiço indisponível (creds): {escopo[0]} - {escopo[1]}\n")
   sys.exit()

# Autenticar e acessar a planilha
cliente = gspread.authorize(credenciais)

# Abra a planilha
planilha = cliente.open_by_url(os.getenv('BANCO_1_URL'))
if not planilha:
   funcoes.banco_indisponivel("", os.getenv('BANCO_1_URL'), "")
   sys.exit()

titulo = planilha.title
sheets = planilha.worksheets()
#=============================================================================================
# variáveis para montagem da planilha
#=============================================================================================
eventos_sensibilizacao = {}
eventos_prospeccao = {}
eventos_qualificacao = {}
pessoas_sensibilizadas = {}
pessoas_prospectadas = {}
pessoas_qualificadas = {}
#=============================================================================================
# Eventos
#=============================================================================================
for sheet in sheets:
    # aba conteúdo total
    aba_total = sheet.get_all_values()
    # Cabeçalhos: primeira linha [0]
    cabecalho = aba_total[0]
    # Dados: segunda linha em diante [1:]
    linhas = aba_total[1:]
    # DataFrama definido com 'cabecalho' e 'linhas'
    dados = pd.DataFrame(linhas, columns=cabecalho)
    if sheet.title == 'Dados de Inscrições em Eventos':
       # Convertendo a coluna 'Data do Evento' para o formato de data, caso não esteja
       dados['Data do Evento'] = pd.to_datetime(dados['Data do Evento'], errors='coerce', dayfirst=True)
       # Removendo itens duplicados
       resultado1 = dados[['Data do Evento', 'Evento']].drop_duplicates()
       # Contabilizando os eventos por data (mantendo a descrição do evento e data) e definindo a coluna 'Ano'
       resultado2 = resultado1.groupby([resultado1['Data do Evento'].dt.year.rename('Ano'), 'Evento']).size().reset_index(name='Total de Eventos')
       # Agrupando por ano e somando a quantidade total de eventos por ano
       resultado3 = resultado2.groupby('Ano')['Total de Eventos'].sum().reset_index(name='Total de Eventos por Ano')
       funcoes.listar_tabela('Eventos: sensibilização', ['Ano', 'Eventos'], resultado3, 'Total de Eventos por Ano')
       for index, row in resultado3.iterrows():
           ano = row['Ano']
           eventos_sensibilizacao[ano] = eventos_sensibilizacao.get(ano, 0) + row['Total de Eventos por Ano']
    if sheet.title == 'Dados de Prospecção e Qualificação':
       # Convertendo a coluna 'Data do Evento' para o formato de data, caso não esteja
       dados['Data do Evento:'] = pd.to_datetime(dados['Data do Evento:'], errors='coerce', dayfirst=True)

       # Extraindo o ano da coluna 'Data do Evento:'
       dados['Ano do Evento'] = dados['Data do Evento:'].dt.year

       # Filtrando para manter apenas os eventos do tipo 'Prospecção' (exclusivos)
       df_prospec = dados[dados['Tipo do Evento:'] == 'Prospecção'].drop_duplicates(subset=['Nome do Evento:', 'Data do Evento:'])

       # Contabilizando o número de eventos 'Prospecção' por ano
       resultado = df_prospec.groupby(['Ano do Evento']).size().reset_index(name='Quantidade')
       funcoes.listar_tabela('Eventos: prospecção', None, resultado, 'Quantidade')
       for index, row in resultado.iterrows():
           ano = row['Ano do Evento']
           eventos_prospeccao[ano] = eventos_prospeccao.get(ano, 0) + row['Quantidade']

       # Filtrando para manter apenas os eventos do tipo 'Qualificação' (exclusivos)
       df_prospec = dados[dados['Tipo do Evento:'] == 'Qualificação'].drop_duplicates(subset=['Nome do Evento:', 'Data do Evento:'])

       # Contabilizando o número de eventos 'Prospecção' por ano
       resultado = df_prospec.groupby(['Ano do Evento']).size().reset_index(name='Quantidade')
       funcoes.listar_tabela('Eventos: qualificação', None, resultado, 'Quantidade')
       for index, row in resultado.iterrows():
           ano = row['Ano do Evento']
           eventos_qualificacao[ano] = eventos_qualificacao.get(ano, 0) + row['Quantidade']

#=============================================================================================
# Pessoas
#=============================================================================================
dados_satisfacao = []
dados_prospeccao = []
for sheet in sheets:
    # conteúdo da aba
    aba_total = sheet.get_all_values()
    # Cabeçalhos: primeira linha [0]
    cabecalho = aba_total[0]
    # Dados: segunda linha em diante [1:]
    linhas = aba_total[1:]
    # DataFrama definido com 'cabecalho' e 'linhas'
    dados = pd.DataFrame(linhas, columns=cabecalho)
    if sheet.title == 'Dados de Inscrições em Eventos':
       # Obter as combinações únicas de 'Data do Evento' e 'Pessoas'
       dados['Data do Evento'] = pd.to_datetime(dados['Data do Evento'], errors='coerce', dayfirst=True)
       # Verificar se há valores nulos em 'Data do Evento' após a conversão
       if dados['Data do Evento'].isnull().any():
          print("Atenção: Existem valores inválidos em 'Data do Evento'")       
       # Remover linhas onde 'Data do Evento' ou 'Pessoas' são vazias
       dados = dados.dropna(subset=['Data do Evento', 'Pessoas'])
       # Remover entradas onde 'Pessoas' são strings vazias (se necessário)
       dados = dados[dados['Pessoas'].str.strip() != '']
       resultado1 = dados[['Data do Evento', 'Pessoas']].drop_duplicates(subset=['Data do Evento', 'Pessoas'])
       resultado1['Ano'] = resultado1['Data do Evento'].dt.year
       # Remover entradas onde 'Ano' são strings vazias (se necessário)
       resultado1 = resultado1.dropna(subset=['Ano'])
       resultado1 = resultado1[resultado1['Ano'] != 0]
       if resultado1['Ano'].isnull().any():
          print("Atenção: Existem valores inválidos em 'Ano'")       
       resultado2 = resultado1.groupby('Ano')['Pessoas'].count().to_dict()
       # Converter o dicionário resultado3 para uma lista de listas no formato [ano, quantidade]
       resultado3 = [[ano, quantidade] for ano, quantidade in resultado2.items()]
       for ano, quantidade in resultado3:
           ano = int(ano)
           quantidade = int(quantidade)
           pessoas_sensibilizadas[ano] = pessoas_sensibilizadas.get(ano, 0) + quantidade
       funcoes.listar_tabela('Pessoas: sensibilizadas', ['Ano', 'Quantidade'], resultado3, 'Quantidade')

    if sheet.title == 'Dados de Satisfação em Eventos':
       dados_satisfacao = dados
    if sheet.title == 'Dados de Prospecção e Qualificação':
       dados_prospeccao = dados

if not dados_satisfacao.empty and not dados_prospeccao.empty:
    #dicionario_eventos = []
    dicionario_eventos = {}
    for index, linha in dados_prospeccao.iterrows():
        data_evento = linha.get('Data do Evento:', '').strip() # Usando .get() para evitar KeyError
        nome_evento = linha.get('Nome do Evento:', '').strip()
        tipo_evento = linha.get('Tipo do Evento:', '').strip()
        #if data_evento and nome_evento and tipo_evento == 'Prospecção':
        if data_evento and nome_evento:
           dicionario_eventos[(data_evento, nome_evento, tipo_evento)] = True 
    #=========================================================================================================
    # Pessoas: prospecção
    #=========================================================================================================
    pessoas_prospeccao_eventos_encontrados = []
    pessoas_prospeccao_resultado_anual = {}  # Dicionário para armazenar a contagem de eventos por ano
    pessoas_prospeccao_ano_contagem = {}
    for index, linha in dados_satisfacao.iterrows():
        data_satisfacao = linha.get('Data do evento', '').strip()
        evento_satisfacao = linha.get('Evento', '').strip()
        email = linha.get('E-mail', '').strip()
        if (data_satisfacao, evento_satisfacao, "Prospecção") in dicionario_eventos:
           pessoas_prospeccao_eventos_encontrados.append((email, data_satisfacao, evento_satisfacao))
           try:
                # Converte a string da data para um objeto datetime e extrai o ano
                ano = pd.to_datetime(data_satisfacao, dayfirst=True).year
                # Incrementa a contagem para aquele ano
                pessoas_prospeccao_ano_contagem[ano] = pessoas_prospeccao_ano_contagem.get(ano, 0) + 1
           except ValueError:
                print(f"Aviso: Não foi possível extrair o ano da data: {data_satisfacao}")
    pessoas_prospeccao_resultado_anual = [[ano, count] for ano, count in sorted(pessoas_prospeccao_ano_contagem.items())]
    for ano, quantidade in pessoas_prospeccao_resultado_anual:
        ano = int(ano)
        quantidade = int(quantidade)
        pessoas_prospectadas[ano] = pessoas_prospectadas.get(ano, 0) + quantidade
    funcoes.listar_tabela('Pessoas: prospectadas', ['Ano', 'Quantidade'], pessoas_prospeccao_resultado_anual, 'Quantidade')
    #=========================================================================================================
    # Pessoas: qualificação
    #=========================================================================================================
    pessoas_qualificacao_eventos_encontrados = []
    pessoas_qualificacao_resultado_anual = {}  # Dicionário para armazenar a contagem de eventos por ano
    pessoas_qualificacao_ano_contagem = {}
    for index, linha in dados_satisfacao.iterrows():
        data_satisfacao = linha.get('Data do evento', '').strip()
        evento_satisfacao = linha.get('Evento', '').strip()
        email = linha.get('E-mail', '').strip()
        if (data_satisfacao, evento_satisfacao, "Qualificação") in dicionario_eventos:
           pessoas_qualificacao_eventos_encontrados.append((email, data_satisfacao, evento_satisfacao))
           try:
                # Converte a string da data para um objeto datetime e extrai o ano
                ano = pd.to_datetime(data_satisfacao, dayfirst=True).year
                # Incrementa a contagem para aquele ano
                pessoas_qualificacao_ano_contagem[ano] = pessoas_qualificacao_ano_contagem.get(ano, 0) + 1
           except ValueError:
                print(f"Aviso: Não foi possível extrair o ano da data: {data_satisfacao}")
    pessoas_qualificacao_resultado_anual = [[ano, count] for ano, count in sorted(pessoas_qualificacao_ano_contagem.items())]
    for ano, quantidade in pessoas_qualificacao_resultado_anual:
        ano = int(ano)
        quantidade = int(quantidade)
        pessoas_qualificadas[ano] = pessoas_qualificadas.get(ano, 0) + quantidade
    funcoes.listar_tabela('Pessoas: qualificadas', ['Ano', 'Quantidade'], pessoas_qualificacao_resultado_anual,'Quantidade')

#=============================================================================================
# Planilha com resultados
#=============================================================================================
# Cabeçalho da tabela
cabecalho_planilha = [
    ['Ano', 'Eventos: Sensibilização', 'Eventos: Prospecção', 'Eventos: Qualificação',
     'Pessoas: Sensibilizadas', 'Pessoas: Prospectadas', 'Pessoas: Qualificadas']
]
# Lista com as contagens por ano:
anos = sorted(set(eventos_sensibilizacao.keys()).union(eventos_prospeccao.keys(), eventos_qualificacao.keys()))
dados_planilha = []
for ano in anos:
    eventos_sens = eventos_sensibilizacao.get(ano, 0)
    eventos_prosp = eventos_prospeccao.get(ano, 0)
    eventos_qual = eventos_qualificacao.get(ano, 0)
    pessoas_sens = pessoas_sensibilizadas.get(ano, 0)
    pessoas_prosp = pessoas_prospectadas.get(ano, 0)
    pessoas_qual = pessoas_qualificadas.get(ano, 0)
    dados_planilha.append([int(ano), int(eventos_sens), int(eventos_prosp), int(eventos_qual),
                           int(pessoas_sens), int(pessoas_prosp), int(pessoas_qual)])

# Capture o retorno da função driver_conexao
service_drive, service_sheets = funcoes.driver_conexao()

# id da URL da pasta compartilhada
url_id = funcoes.link_id(os.getenv('PASTA_COMPARTILHADA'))

subpasta_id = ''

print(f"\nPreparando planilha dos resultados...")

try:
    # temporário para apagar o conteúdo criado (pasta e planilhas)
    #funcoes.apagar_pasta_arquivo(service_drive, None, 'pasta_13-38-01', url_id)
    #funcoes.apagar_pasta_arquivo(service_drive, None, 'pasta_teste1', url_id)
    funcoes.apagar_pasta_arquivo(service_drive, None, 'indicadores', url_id)

    # Criar a subpasta dentro da pasta compartilhada
    if os.getenv('SUB_PASTA'):
       subpasta_id = funcoes.criar_pasta(service_drive, os.getenv('SUB_PASTA'), url_id)
    
    # nome da planilha Google Sheets
    planilha_nome = os.getenv('PLANILHA')

    # criar planilha na subpasta
    planilha_id = funcoes.criar_planilha(service_drive, service_sheets, planilha_nome, subpasta_id)

    # Permissões da planilha
    funcoes.permissoes_pasta_arquivo(service_drive, planilha_id, 'anyone', 'writer')
    # Preparando cabecalho
    cabecalho_intervalo = funcoes.planilha_celulas_intervalo(
        os.getenv('planilha_coluna_inicial','a'),
        os.getenv('planilha_linha_inicial','1'),
        cabecalho_planilha,'c' # c-cabeçalho d-dados
    )
    dados_intervalo = funcoes.planilha_celulas_intervalo(
        os.getenv('planilha_coluna_inicial','a'),
        os.getenv('planilha_linha_inicial','3'),
        dados_planilha,'d' # c-cabeçalho d-dados
    )

    funcoes.planilha_dados(service_sheets, planilha_id, cabecalho_intervalo, cabecalho_planilha)
    funcoes.planilha_dados(service_sheets, planilha_id, dados_intervalo, dados_planilha)

    funcoes.planilha_formatacao(service_sheets, planilha_id, 'Banco 1', cabecalho_intervalo, dados_intervalo)
        
    # todos os métodos abaixo estão funcionando: 21/05/2025 ás 13:29
    #funcoes.apagar_pasta_arquivo(service_drive, subpasta_id, None, url_id)
    #funcoes.apagar_pasta_arquivo(service_drive, planilha_id, None, url_id)
    #funcoes.apagar_pasta_arquivo(service_drive, None, subpasta_nome, url_id)
    #funcoes.apagar_pasta_arquivo(service_drive, None, planilha_nome, subpasta_id)

    # Após a criação e formatação da planilha
    planilha_url = f"https://docs.google.com/spreadsheets/d/{planilha_id}/edit"
    print(f"\nLink da planilha: {planilha_url}")

    print("\nFim.\n")

except HttpError as error:
    print(f"\nOcorreu um erro: {error}\n")
