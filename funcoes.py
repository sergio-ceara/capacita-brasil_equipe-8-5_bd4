import os
import re
import sys
import socket
import gspread
import pandas as pd
from tabulate import tabulate
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import range_boundaries

# Variáveis constantes (utilizadas de forma global no programa)
GOOGLE_DRIVE_SCOPE       = 'https://www.googleapis.com/auth/drive'
GOOGLE_SHEETS_SCOPE      = 'https://www.googleapis.com/auth/spreadsheets'
GOOGLE_DNS_SERVER        = "8.8.8.8"
GOOGLE_DNS_PORT          = 53
CONNECTION_TIMEOUT       = 5
DEFAULT_SHEET_ID         = 0  # Geralmente o ID da primeira aba
DEFAULT_BACKGROUND_COLOR = {'red': 0.85, 'green': 0.85, 'blue': 0.85}

# Verificar conexão com a internet
def verificar_conexao(host=GOOGLE_DNS_SERVER, port=GOOGLE_DNS_PORT, timeout=CONNECTION_TIMEOUT):
    try:
        socket.create_connection((host, port), timeout=timeout)
        print("\nConectado a internet.")
        return True
    except OSError as e:
        print(f"Erro de conexão com a internet: {e}")
        return False

# Inicia os serviços da API do Google Drive, Google Sheets e autentica o cliente gspread.
# Retorna uma tupla contendo as instâncias de serviço do Drive, Sheets e o cliente gspread.
def conectar_google_apis():
    SCOPES = [GOOGLE_DRIVE_SCOPE, GOOGLE_SHEETS_SCOPE, "https://spreadsheets.google.com/feeds"]
    creds_path = os.getenv('GOOGLE_CREDS_JSON_PATH')
    # Salto de linha para as mensagens abaixo
    print("")
    if not creds_path:
        print("Arquivo de credenciais não especificado na variável de ambiente 'GOOGLE_CREDS_JSON_PATH'.")
        sys.exit(1)

    try:
        # Autenticação para googleapiclient (Drive e Sheets)
        creds          = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        service_drive  = build('drive', 'v3', credentials=creds)
        service_sheets = build('sheets', 'v4', credentials=creds)
        # Autenticação para gspread
        gs_credentials = ServiceAccountCredentials.from_json_keyfile_name(creds_path, SCOPES)
        if not gs_credentials:
            print("Serviço gspread indisponível (credenciais inválidas).")
            sys.exit(1)
        gs_client = gspread.authorize(gs_credentials)
        print("Serviços do Google Drive, Sheets e cliente gspread ativos.")
        return service_drive, service_sheets, gs_client
    except Exception as e:
        print(f"Falha ao carregar credenciais ou construir serviços: {e}")
        sys.exit(1)

# Extrai o ID de um link de arquivo ou pasta do Google Drive.
def link_id(service_drive, url: str):
    url_id = re.search(r"(?:folders|file)/([a-zA-Z0-9_-]+)", url)
    if url_id:
        id = url_id.group(1)
        informacoes = informacoes_driver(service_drive, id)
        print(f"\nNome da pasta compartilhada: {informacoes[0]}")
        print(f"link: {url}")
        return id
    else:
        print(f"ID não encontrado na URL: {url}")
        return ""

# Registra uma mensagem de banco de dados indisponível.
def banco_indisponivel(banco: str, url: str):
    mensagem1 = f"Banco: {banco}"
    if url:
        mensagem2 = f"URL: {url}"
        mensagem3 = "Situação: indisponível."
    else:
        mensagem2 = f"URL: {url}"
        mensagem3 = "Situação: URL vazia."
    mensagem_final = f"{mensagem1}, {mensagem2}, {mensagem3}"
    print(mensagem_final)

# Prepara os dados para exibição em uma tabela, convertendo-os para um DataFrame.
# O parâmetro 'registros' pode ser recebido nos seguintes tipos: list, dict, pd.DataFrame, pd.Series
# O parâmetro 'cabecalho' pode ser recebido nos seguinte tipo: list
def preparar_dados_para_tabela(registros, cabecalho: list = None):
    df = pd.DataFrame()
    if isinstance(registros, dict):
        try:
            df = pd.DataFrame(registros)
        except ValueError as e:
            print(f"Erro ao converter dicionário em DataFrame: {e}. Verifique se todas as listas de valores têm o mesmo comprimento.")
            raise
    elif isinstance(registros, list):
        if not cabecalho:
            print("Cabeçalho necessário para registros do tipo lista. Tentando inferir...")
            if registros and all(isinstance(r, list) for r in registros):
                cabecalho = [f"Coluna {i+1}" for i in range(len(registros[0]))]
            else:
                raise ValueError("Formato de lista de registros inválido. Espera-se uma lista de listas.")
        df = pd.DataFrame(registros, columns=cabecalho)
    elif isinstance(registros, (pd.DataFrame, pd.Series)):
        df = pd.DataFrame(registros) if isinstance(registros, pd.Series) else registros
    else:
        raise ValueError(f"Formato de 'registros' {type(registros)} não suportado.")

    if cabecalho and list(df.columns) != cabecalho:
        df.columns = cabecalho
    
    return df

# Exibe um DataFrame Pandas em uma tabela formatada no console, incluindo totais.
def exibir_tabela_formatada(titulo: str, df: pd.DataFrame, coluna_numerica: str = None):
    print(f"\n{titulo}")

    if df.empty:
        print("Nenhum registro para exibir.")
        return

    quantidade_registros   = len(df)
    soma_coluna_numerica   = 0
    coluna_soma_encontrada = False

    if coluna_numerica:
        if coluna_numerica in df.columns:
            try:
                df[coluna_numerica] = pd.to_numeric(df[coluna_numerica], errors='coerce')
                soma_coluna_numerica = df[coluna_numerica].sum()
                coluna_soma_encontrada = True
            except Exception as e:
                print(f"Não foi possível somar a coluna '{coluna_numerica}': {e}. Ignorando soma para esta coluna.")
        else:
            print(f"A coluna '{coluna_numerica}' não foi encontrada nos dados. Nenhuma soma será exibida para ela.")
    else:
        for col_name in df.columns:
            if pd.api.types.is_numeric_dtype(df[col_name]):
                coluna_numerica = col_name
                soma_coluna_numerica = df[col_name].sum()
                coluna_soma_encontrada = True
                break

    ultimo_registro = [""] * len(df.columns)
    ultimo_registro[0] = f"{quantidade_registros} itens"

    if coluna_soma_encontrada and coluna_numerica:
        try:
            indice_coluna_numerica = df.columns.get_loc(coluna_numerica)
            ultimo_registro[indice_coluna_numerica] = f"{soma_coluna_numerica:>10.2f}"
        except KeyError:
            print(f"Coluna numérica '{coluna_numerica}' não encontrada para posicionamento da soma.")

    df_com_total = pd.concat([df, pd.DataFrame([ultimo_registro], columns=df.columns)], ignore_index=True)
    
    print(tabulate(df_com_total, headers="keys", tablefmt="grid"))

# Carrega dados de uma aba específica de uma planilha, realiza transformações e validações.
# Retorna um DataFrame com os dados transformados e a lista de colunas finais.
def carregar_dados_planilha(planilha, aba: str):
    try:
        sheet = planilha.worksheet(aba)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Aba '{aba}' não encontrada na planilha.")
        sys.exit(1)

    dados = pd.DataFrame(sheet.get_all_records())
    # A condição "IF" é utiliza para acrescentar outras abas e seus processamentos.
    if aba == 'Dados Seleção':
       colunas_necessarias = ['Ano', 'Nome', 'Contrato', 'Cidade', 'Estado', 'Área']
       if not all(col in dados.columns for col in colunas_necessarias):
          missing_cols = [col for col in colunas_necessarias if col not in dados.columns]
          print(f"Colunas obrigatórias faltando na aba '{aba}': {', '.join(missing_cols)}")
          sys.exit(1)
        
       dados = dados[colunas_necessarias].copy()
       # Acrescentar uma coluna fixa 'País' com conteúdo "Brasil"
       dados['País'] = 'Brasil'
       # Transforma o conteúdo da coluna 'Contrato' de 'Sim' para 'Aprovado' e 'Não' para 'Inscrito'.
       dados['Contrato'] = dados['Contrato'].replace({'Sim': 'Aprovado', 'Não': 'Inscrito'})
       # Altera o nome das colunas para ficar compatível com o dashboard criado.
       dados = dados.rename(columns={
            'Nome': 'identificação',
            'Contrato': 'status',
            'Cidade': 'cidade',
            'Estado': 'UF',
            'Área': 'area'
        })
       # Define os títulos das colunas com os novos nomes
       colunas_final = ['Ano', 'identificação', 'status', 'cidade', 'UF', 'País', 'area']
       # Atribui esses novos nomes 
       dados = dados[colunas_final]
       # Contabiliza registro duplicados 
       repetidos = dados.duplicated().sum()
       if repetidos > 0:
          print(f"Quantidade de registros repetidos encontrados: {repetidos}")
       else:
           print("\nNenhum registro repetido encontrado.")

    return dados, colunas_final

# Prepara os intervalos de células para o cabeçalho e os dados na planilha.
# Retorna uma tupla contendo as strings dos intervalos do cabeçalho e dos dados.
def preparar_intervalos(cabecalho: list, dados: pd.DataFrame):
    col_ini   = os.getenv('planilha_coluna_inicial', 'A').upper()
    linha_ini = os.getenv('planilha_linha_inicial', '1')

    try:
        linha_ini = int(linha_ini)
    except ValueError:
        print(f"planilha_linha_inicial '{linha_ini}' inválida. Usando '1'.")
        linha_ini = 1

    # Para o cabeçalho, a linha inicial é a definida
    intervalo_cabecalho = planilha_celulas_intervalo(col_ini, linha_ini, [cabecalho], 'c')
    
    # Para os dados, a linha inicial é a linha do cabeçalho + 1 (onde os dados efetivamente começam)
    linha_dados_ini = linha_ini + 1
    intervalo_dados = planilha_celulas_intervalo(col_ini, linha_dados_ini, dados.values.tolist(), 'd')

    return intervalo_cabecalho, intervalo_dados

# Verifica se uma pasta existe no Google Drive pelo nome (pasta_nome) e local onde pesquisar (parent_id: pasta pai)
def pasta_existe(service_drive, pasta_nome: str, parent_id: str = None):
    query = f"name = '{pasta_nome}' and mimeType = 'application/vnd.google-apps.folder'"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    
    try:
        results = service_drive.files().list(q=query, fields="files(id)").execute()
        files = results.get('files', [])
        if files:
            print(f"\nPasta '{pasta_nome}' já existe.")
            print(f"https://drive.google.com/drive/folders/{files[0]['id']}")
            return files[0]['id']
        print(f"\nPasta '{pasta_nome}' não encontrada.")
        return None
    except HttpError as e:
        print(f"Erro ao verificar a existência da pasta '{pasta_nome}': {e}")
        return None

# Verifica se uma planilha existe no Google Drive pelo nome (planilha_nome) e local onde pesquisar (parent_id: pasta pai)
def planilha_existe(service_drive, planilha_nome: str, parent_id: str):
    query = f"name = '{planilha_nome}' and mimeType = 'application/vnd.google-apps.spreadsheet'"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    
    try:
        results = service_drive.files().list(q=query, fields="files(id)").execute()
        files = results.get('files', [])
        if files:
            print(f"\nPlanilha '{planilha_nome}' já existe.")
            print(f"https://docs.google.com/spreadsheets/d/{files[0]['id']}")
            return files[0]['id']
        print(f"Planilha '{planilha_nome}' não encontrada.")
        return None
    except HttpError as e:
        print(f"Erro ao verificar a existência da planilha '{planilha_nome}': {e}")
        return None

# Cria uma pasta no Google Drive, com opção de definir uma pasta pai.
# Se a pasta já existir, retorna o ID da pasta existente.
# Atribui permissões 'anyone' e 'writer' nas duas situações.
def criar_pasta(service_drive, pasta_nome: str, parent_id: str = None):
    pasta_id = pasta_existe(service_drive, pasta_nome, parent_id)
    if pasta_id:
        #print(f"\nPasta '{pasta_nome}' já existe.")
        permissoes_pasta_arquivo(service_drive, pasta_id, 'anyone', 'writer')
        return pasta_id

    file_metadata = {
        'name': pasta_nome,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    if parent_id:
        file_metadata['parents'] = [parent_id]

    try:
        pasta = service_drive.files().create(body=file_metadata, fields='id').execute()
        pasta_id = pasta['id']
        print(f"Criando pasta '{pasta_nome}' com ID: {pasta_id}...")
        permissoes_pasta_arquivo(service_drive, pasta_id, 'anyone', 'writer')
        return pasta_id
    except HttpError as e:
        print(f"Erro ao criar a pasta '{pasta_nome}': {e}")
        raise

# Cria uma planilha no Google Sheets, com opção de definir uma pasta pai.
# Se a planilha já existir, retorna o ID da planilha existente.
def criar_planilha(service_drive, service_sheets, planilha_nome: str, parent_id: str = None):
    planilha_id = planilha_existe(service_drive, planilha_nome, parent_id)
    if planilha_id:
        return planilha_id

    body = {
        'properties': {'title': planilha_nome}
    }
    try:
        planilha = service_sheets.spreadsheets().create(body=body, fields='spreadsheetId').execute()
        planilha_id = planilha['spreadsheetId']
        print(f"Criando planilha '{planilha_nome}' com ID: {planilha_id}...")

        if parent_id:
            service_drive.files().update(
                fileId=planilha_id,
                addParents=parent_id,
                fields='parents'
            ).execute()
            print(f"Planilha '{planilha_nome}' movida para a pasta pai '{parent_id}'.")
        
        return planilha_id
    except HttpError as e:
        print(f"Erro ao criar a planilha '{planilha_nome}': {e}")
        raise

# Atribui permissões por tipo e função a uma pasta ou arquivo no Google Drive.
def permissoes_pasta_arquivo(service_drive, item_id: str, tipo: str, funcao: str):
    informacoes = informacoes_driver(service_drive, item_id)
    item_nome = informacoes[0]
    item_tipo_mime = informacoes[1]

    if 'folder' in item_tipo_mime:
        tipo_item_str = 'Pasta'
    else:
        tipo_item_str = 'Arquivo'

    mensagem = f"{tipo_item_str}: '{item_nome}' - Atribuindo permissões de '{tipo}' com função de '{funcao}'."
    permissao = {
        'type': tipo,
        'role': funcao
    }

    try:
        service_drive.permissions().create(
            fileId=item_id,
            body=permissao
        ).execute()
        print(f"\nPermissões atribuídas com sucesso para {mensagem}")
    except HttpError as e:
        print(f"\nFalha ao atribuir permissões para {mensagem}: {e}")

# Retorna o nome e o tipo MIME de um item (arquivo ou pasta) no Google Drive.
def informacoes_driver(service_drive, item_id: str):
    try:
        conteudo = service_drive.files().get(fileId=item_id, fields='name,mimeType').execute()
        conteudo_nome = conteudo.get('name', '')
        conteudo_tipo = conteudo.get('mimeType', '')
        return conteudo_nome, conteudo_tipo
    except HttpError as e:
        print(f"Erro ao obter informações para o item com ID '{item_id}': {e}")
        return "", ""

# Grava informações em um intervalo específico de uma planilha.
def planilha_dados(service_sheets, planilha_id: str, intervalo: str, dados_para_gravar: list):
    try:
        request = service_sheets.spreadsheets().values().update(
            spreadsheetId=planilha_id,
            range=intervalo,
            valueInputOption='RAW',
            body={'values': dados_para_gravar}
        )
        response = request.execute()
        print(f"\nDados inseridos com sucesso no intervalo: {intervalo}")
        return response
    except HttpError as e:
        print(f"\nErro ao inserir dados no intervalo '{intervalo}': {e}")
        return None

# Apaga uma pasta ou arquivo pelo ID ou nome fornecido.
def apagar_pasta_arquivo(service_drive, item_id: str = None, item_nome: str = None, parent_id: str = None):
    if not item_id and not item_nome:
        print("A função 'apagar_pasta_arquivo' precisa de um 'item_id' ou 'item_nome' para continuar.")
        return

    if item_nome:
        found_id = pasta_existe(service_drive, item_nome, parent_id)
        if not found_id:
            found_id = planilha_existe(service_drive, item_nome, parent_id)
        
        if not found_id:
            print(f"Não foi encontrado 'pasta' ou 'planilha' com o nome '{item_nome}'.")
            return
        item_id = found_id

    informacoes = informacoes_driver(service_drive, item_id)
    item_nome_real = informacoes[0]
    item_tipo_mime = informacoes[1]
    # Se for uma pasta
    if 'folder' in item_tipo_mime:
        query = f"'{item_id}' in parents"
        try:
            results = service_drive.files().list(q=query, fields="files(id, name)").execute()
            files_in_folder = results.get('files', [])

            if not files_in_folder:
                service_drive.files().delete(fileId=item_id).execute()
                print(f"Pasta '{item_nome_real}' foi removida com sucesso!")
            else:
                print(f"A pasta '{item_nome_real}' contém {len(files_in_folder)} arquivo(s).")
                resposta = input("Digite 'sim' para apagar os arquivos dentro dela e a pasta: ").strip().lower()
                if resposta == 'sim':
                    for arquivo in files_in_folder:
                        try:
                            service_drive.files().delete(fileId=arquivo['id']).execute()
                            print(f"Arquivo '{arquivo['name']}' apagado com sucesso!")
                        except HttpError as error:
                            print(f"Erro ao apagar o arquivo '{arquivo['name']}': {error}")
                    
                    try:
                        service_drive.files().delete(fileId=item_id).execute()
                        print(f"Pasta '{item_nome_real}' foi removida com sucesso, juntamente com seus arquivos!")
                    except HttpError as error:
                        print(f"Erro ao remover a pasta '{item_nome_real}': {error}")
                else:
                    print("A exclusão da pasta foi cancelada.")
        except HttpError as e:
            print(f"Erro ao verificar ou apagar a pasta '{item_nome_real}': {e}")
    # Se for uma arquivo (planilha)
    else:
        try:
            service_drive.files().delete(fileId=item_id).execute()
            print(f"Arquivo '{item_nome_real}' foi removido com sucesso!")
        except HttpError as e:
            print(f"Erro ao remover o arquivo '{item_nome_real}': {e}")

# Obter o Id da aba da planilha pelo nome.
def id_aba_planilha_por_nome(service_sheets, planilha_id: str, aba_nome: str):
    try:
        spreadsheet_metadata = service_sheets.spreadsheets().get(
            spreadsheetId=planilha_id, fields='sheets.properties'
        ).execute()
        for sheet_prop in spreadsheet_metadata.get('sheets', []):
            if sheet_prop['properties']['title'] == aba_nome:
                return sheet_prop['properties']['sheetId']
        print(f"Aba '{aba_nome}' não encontrada. Usando sheetId padrão (0).")
        return DEFAULT_SHEET_ID
    except HttpError as e:
        print(f"Erro ao obter metadados da planilha para encontrar a aba '{aba_nome}': {e}")
        return DEFAULT_SHEET_ID

# Converte uma string de intervalo (ex: 'B3:H9') para (min_row, max_row, min_col, max_col), utilizando o pacote 'openpyxl'.
def celula_intervalo_para_linhas_colunas(intervalo_str: str):
    min_col, min_row, max_col, max_row = range_boundaries(intervalo_str)
    return min_row -1, max_row, min_col -1, max_col # Ajustar para 0-index

# Planilha: formatação: remover linhas de grade
def formatar_remover_linhas_grade(sheet_id: int):
    return {
        'updateSheetProperties': {
            'properties': {
                'sheetId': sheet_id,
                'gridProperties': {
                    'hideGridlines': True
                }
            },
            'fields': 'gridProperties.hideGridlines'
        }
    }

# Planilha: formatação: cor de fundo do cabeçalho
def formatar_fundo_cabecalho(sheet_id: int, cabecalho_intervalo: str):
    start_row, end_row, start_col, end_col = celula_intervalo_para_linhas_colunas(cabecalho_intervalo)
    return {
        'updateCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': start_row, 
                'endRowIndex': end_row,
                'startColumnIndex': start_col,
                'endColumnIndex': end_col # endColumnIndex é exclusivo
            },
            'fields': 'userEnteredFormat.backgroundColor',
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredFormat': {
                                'backgroundColor': DEFAULT_BACKGROUND_COLOR
                            }
                        }
                    ] * (end_col - start_col) # Usar end_col - start_col para quantidade de colunas
                }
            ]
        }
    }

# Planilha: formatação: bordas num intervalo
def formatar_bordas(sheet_id: int, intervalo: str):
    start_row, end_row, start_col, end_col = celula_intervalo_para_linhas_colunas(intervalo)
    return {
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': start_row,
                'endRowIndex': end_row,
                'startColumnIndex': start_col,
                'endColumnIndex': end_col # endColumnIndex é exclusivo
            },
            'top': {'style': 'SOLID', 'width': 1},
            'bottom': {'style': 'SOLID', 'width': 1},
            'left': {'style': 'SOLID', 'width': 1},
            'right': {'style': 'SOLID', 'width': 1},
            'innerHorizontal': {'style': 'SOLID', 'width': 1},
            'innerVertical': {'style': 'SOLID', 'width': 1}
        }
    }

# Planilha: formatação: renomear aba
def formatar_renomear_aba(sheet_id: int, nova_aba_nome: str):
    return {
        'updateSheetProperties': {
            'properties': {
                'sheetId': sheet_id,
                'title': nova_aba_nome
            },
            'fields': 'title'
        }
    }

# Planilha: formatação: centralizar conteúdo de um intervalo
def formatar_centralizar_conteudo(sheet_id: int, intervalo: str):
    start_row, end_row, start_col, end_col = celula_intervalo_para_linhas_colunas(intervalo)
    return {
        'updateCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': start_row,
                'endRowIndex': end_row,
                'startColumnIndex': start_col,
                'endColumnIndex': end_col # endColumnIndex é exclusivo
            },
            'fields': 'userEnteredFormat.horizontalAlignment',
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredFormat': {
                                'horizontalAlignment': 'CENTER'
                            }
                        }
                    ] * (end_col - start_col)
                }
            ]
        }
    }

# Planilha: formatação: autoajuste de colunas de um intervalo
def formatar_auto_ajustar_colunas(sheet_id: int, start_col: int, end_col: int):
    return {
        'autoResizeDimensions': {
            'dimensions': {
                'sheetId': sheet_id,
                'dimension': 'COLUMNS',
                'startIndex': start_col,
                'endIndex': end_col # endIndex é exclusivo
            }
        }
    }

# Aplica formatações desejadas a uma planilha do Google Sheets.
def aplicar_formatacoes_planilha(service_sheets, planilha_id: str, aba_nome: str, cabecalho_intervalo: str, dados_intervalo: str):
    # Variável para acrescentar as requisições de formatações desejadas
    requisicao = []
    sheet_id = id_aba_planilha_por_nome(service_sheets, planilha_id, aba_nome)

    _, _, start_col_cabecalho, _ = celula_intervalo_para_linhas_colunas(cabecalho_intervalo)
    _, _, _, end_col_dados       = celula_intervalo_para_linhas_colunas(dados_intervalo)

    requisicao.append(formatar_remover_linhas_grade(sheet_id))
    requisicao.append(formatar_fundo_cabecalho(sheet_id, cabecalho_intervalo))
    requisicao.append(formatar_bordas(sheet_id, cabecalho_intervalo))
    requisicao.append(formatar_bordas(sheet_id, dados_intervalo))
    requisicao.append(formatar_renomear_aba(sheet_id, aba_nome))
    requisicao.append(formatar_centralizar_conteudo(sheet_id, cabecalho_intervalo))
    requisicao.append(formatar_auto_ajustar_colunas(sheet_id, start_col_cabecalho, end_col_dados))

    try:
        body = {'requests': requisicao}
        response = service_sheets.spreadsheets().batchUpdate(
            spreadsheetId=planilha_id,
            body=body
        ).execute()
        print("\nFormatação aplicada com sucesso!")
        return response
    except HttpError as e:
        print(f"\nErro ao aplicar formatação: {e}")
        return None

# Constrói uma string de intervalo de células do Google Sheets.
# Retorna a string do intervalo de células (ex: 'A1:C1').
def planilha_celulas_intervalo(letra_inicial: str, linha_inicial: int, conteudo: list, tipo: str):
    if not conteudo or not isinstance(conteudo, list) or not conteudo[0]:
        print("Função 'planilha_celulas_intervalo': parâmetro 'conteudo' incorreto ou vazio.")
        raise ValueError("Parâmetro 'conteudo' inválido.")

    num_colunas = len(conteudo[0])
    num_linhas  = len(conteudo)

    coluna_inicial_numero = column_index_from_string(letra_inicial.upper())
    ultima_coluna_numero  = coluna_inicial_numero + num_colunas - 1
    ultima_coluna_letra   = get_column_letter(ultima_coluna_numero)

    if tipo.upper() == 'C':
        # Intervalo para o cabeçalho (apenas uma linha)
        intervalo = f'{letra_inicial.upper()}{linha_inicial}:{ultima_coluna_letra}{linha_inicial}'
    elif tipo.upper() == 'D':
        # Intervalo para os dados (começa na linha seguinte ao cabeçalho ou linha_inicial se sem cabeçalho)
        linha_final = linha_inicial + num_linhas - 1
        intervalo = f'{letra_inicial.upper()}{linha_inicial}:{ultima_coluna_letra}{linha_final}'
    else:
        print(f"Função 'planilha_celulas_intervalo': parâmetro 'tipo' inválido '{tipo}'. Use 'C' para Cabeçalho ou 'D' para Dados.")
        raise ValueError("Parâmetro 'tipo' inválido.")
    
    return intervalo