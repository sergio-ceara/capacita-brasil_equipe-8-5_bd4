#==============================================================================================
# Bibliotecas
#==============================================================================================
import os
import sys
# chamada de arquivo com métodos para o processamento
import funcoes
from dotenv import load_dotenv
# capturar uma exceção HttpError
from googleapiclient.errors import HttpError 

#==============================================================================================
# Execução do programa
#==============================================================================================

# Caso não tenha conexão com a internet, o programa é interrompido
if not funcoes.verificar_conexao():
   print(f"\nSem conexão com a internet.\n")
   sys.exit()

# Leitura do arquivo '.env'
load_dotenv()

# Conexão com os serviços
try:
   service_drive, service_sheets, cliente = funcoes.conectar_google_apis()
except SystemExit:
   print("Falha na autenticação. Saindo.")
   sys.exit(1)

# Abra a planilha do Banco 2
planilha = cliente.open_by_url(os.getenv('BANCO_2_URL'))
# Se a planilha não estiver disponível o programa é interrompido
if not planilha:
   funcoes.banco_indisponivel("", os.getenv('BANCO_2_URL'), "")
   sys.exit()

# recebe os dados e cabecalho da leitura da planilha e aba selecionada
dados, cabecalho = funcoes.carregar_dados_planilha(planilha, 'Dados Seleção')

pasta_id    = funcoes.link_id(service_drive, os.getenv('PASTA_COMPARTILHADA'))
subpasta_id = ''
if os.getenv('SUB_PASTA'):
   subpasta_id = funcoes.criar_pasta(service_drive, os.getenv('SUB_PASTA'), pasta_id)

planilha_nome = os.getenv('PLANILHA')
planilha_id = funcoes.criar_planilha(service_drive, service_sheets, planilha_nome, subpasta_id)
#funcoes.permissoes_pasta_arquivo(service_drive, planilha_id, 'anyone', 'writer')

intervalo_cabecalho, intervalo_dados = funcoes.preparar_intervalos(cabecalho, dados)

try:
   funcoes.planilha_dados(service_sheets, planilha_id, intervalo_cabecalho, [cabecalho])
   funcoes.planilha_dados(service_sheets, planilha_id, intervalo_dados, dados.values.tolist())
   funcoes.aplicar_formatacoes_planilha(service_sheets, planilha_id, 'Banco 2', intervalo_cabecalho, intervalo_dados)

   print(f"\nLink da planilha criada: https://docs.google.com/spreadsheets/d/{planilha_id}/edit")
   print("\nFim.\n")

except HttpError as error:
   print(f"\nOcorreu um erro: {error}\n")
