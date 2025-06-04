#==============================================================================================
# Bibliotecas
#==============================================================================================
import os
import sys
# chamada de arquivo com métodos para o processamento
import funcoes
import pandas as pd
from dotenv import load_dotenv
# capturar uma exceção HttpError
from googleapiclient.errors import HttpError 

#==============================================================================================
# Início da execução: carregamento de dados, geração e formatação da planilha
#==============================================================================================

# Verifica conexão com a internet. Encerra o programa se estiver offline.
if not funcoes.verificar_conexao():
   print(f"\nSem conexão com a internet.\n")
   sys.exit()

# Carrega variáveis de ambiente do arquivo .env (URLs, IDs, nomes de planilhas, etc.)
load_dotenv()

# Conecta às APIs do Google Drive, Sheets e gspread usando credenciais do serviço
try:
   service_drive, service_sheets, cliente = funcoes.conectar_google_apis()
except SystemExit:
   print("Falha na autenticação. Saindo.")
   sys.exit(1)

# Abre a planilha principal a partir da URL definida na variável de ambiente BANCO_4_URL
try:
    planilha = cliente.open_by_url(os.getenv('BANCO_4_URL'))
except Exception as e:
    funcoes.banco_indisponivel("[Banco 4] - Agregação de Valor", os.getenv('BANCO_4_URL'))
    print(f"Falha: {str(e)}")
    sys.exit()

# Lê os dados e cabeçalhos das abas "Banco de Consultorias" e "Banco de Mentorias"
cons_dados, cons_cabecalho = funcoes.carregar_dados_planilha(planilha, 'Banco de Consultorias')
ment_dados, ment_cabecalho = funcoes.carregar_dados_planilha(planilha, 'Banco de Mentorias')
# Verifica se ambos os DataFrames estão vazios antes de continuar com a concatenação
if cons_dados.empty and ment_dados.empty:
    print("\nNenhum dado carregado das abas 'Consultoria' e 'Mentoria'. Encerrando.")
    sys.exit(1)

# Junta os dados das duas abas
dados = pd.concat([cons_dados, ment_dados], ignore_index=True)
# Remove linhas totalmente vazias
dados = dados.dropna(how='all')  
print(f"\nDados juntos (mentoria+consultoria): {len(dados)}")

pasta_id    = funcoes.link_id(service_drive, os.getenv('PASTA_COMPARTILHADA'))
subpasta_id = ''
if os.getenv('SUB_PASTA'):
   subpasta_id = funcoes.criar_pasta(service_drive, os.getenv('SUB_PASTA'), pasta_id)
# Valida existência da variável PLANILHA no .env, necessária para nomear a nova planilha
if not os.getenv('PLANILHA'):
    print("Erro: Variável PLANILHA não encontrada no .env")
    sys.exit(1)
planilha_nome = os.getenv('PLANILHA')
planilha_id = funcoes.criar_planilha(service_drive, service_sheets, planilha_nome, subpasta_id)

intervalo_cabecalho, intervalo_dados = funcoes.preparar_intervalos(cons_cabecalho, dados)

try:
   # Insere os dados na nova planilha e aplica formatações (borda, cabeçalho, colunas)
   funcoes.planilha_dados(service_sheets, planilha_id, intervalo_cabecalho, [cons_cabecalho], 's')
   funcoes.planilha_dados(service_sheets, planilha_id, intervalo_dados, dados.values.tolist())
   funcoes.aplicar_formatacoes_planilha(service_sheets, planilha_id, 'Banco 4', intervalo_cabecalho, intervalo_dados)

   print(f"\nLink da planilha criada: https://docs.google.com/spreadsheets/d/{planilha_id}/edit")
   print("\nFim.\n")

except HttpError as error:
   print(f"\nOcorreu um erro: {error}\n")
