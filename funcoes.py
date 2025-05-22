import os
import re
import sys
import socket
import pandas as pd
from tabulate import tabulate
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.service_account import Credentials

# verificar conexão com a internet
def verificar_conexao():
    try:
        # Tenta se conectar ao Google DNS para verificar se há internet
        socket.create_connection(("8.8.8.8", 53), timeout=5)
        return True
    except OSError:
        return False

def driver_conexao():
    # Definindo os escopos necessários para a autenticação
    SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

    # Carregar credenciais de serviço (se você estiver usando uma conta de serviço)
    creds = Credentials.from_service_account_file(
        os.getenv('GOOGLE_CREDS_JSON_PATH'), scopes=SCOPES
    )

    # Criando o cliente para as APIs Drive e Sheets
    service_drive = build('drive', 'v3', credentials=creds)
    service_sheets = build('sheets', 'v4', credentials=creds)
    return service_drive, service_sheets

# Mensagem para banco indisponível
def banco_indisponivel(banco, url, arquivo_saida=None):
    mensagem1 = f"banco: {banco}"
    if url:
       mensagem2 = f"url: {url}"
       mensagem3 = "Situação: indisponível."
    else:
       mensagem2 = f"url: {url}"
       mensagem3 = "Situação: URL vazia."
    if arquivo_saida:
       arquivo_texto_gravar(arquivo_saida,[mensagem1, mensagem2, mensagem3])
    return

def arquivo_texto_gravar(arquivo_saida, vetor_conteudo):
    conteudo = ''
    # Gravando conteúdo no arquivo texto
    # Redirecionar a saída padrão para o arquivo
    sys.stdout = arquivo_saida
    for linha in vetor_conteudo:
        arquivo_saida.write(linha+"\n") # Escreve cada linha no arquivo. O "\n" faz o salto de linha.
        conteudo+=linha+"\n"            # Acumula o conteúdo para exibição posterior
    # Restaurar a saída padrão
    sys.stdout = sys.__stdout__
    # Mostrando o conteúdo no console
    print(conteudo, end='')

def dados_preparar_lista(titulo, cabecalho, chaves, valores):
    registros = [[chave, valor] for chave, valor in zip(chaves, valores)]
    listar_tabela(titulo, cabecalho, registros)

def listar_tabela(titulo, cabecalho, registros, coluna_numerica=0):
    """
    Exibe uma tabela formatada com um título, aceitando diferentes formatos de dados.
    Argumentos:
              titulo (str): O título da tabela.
           cabecalho (str): Título das colunas (Obrigatório para tipo 'list')
          registros (list): Uma lista de dicionários (opção 1) ou uma lista de listas (opção 2).
           coluna_numerica: indicação do nome da coluna que deve ser feito um somatório.
    """
    print(f"\n{titulo}")

    if len(registros) == 0:
        print("Nenhum registro para exibir.")
        return

    if isinstance(registros, dict):
        cabecalho = list(registros.keys())
        #registros = [list(registros.values()) for registro in registros]
        # Usando zip para agrupar os valores de cada chave e formar as linhas corretamente
        registros = list(zip(*registros.values()))  # Transpor as listas para formar as linhas
        # ==============================================================
        quantidade_registros = len(registros)
        soma_colunas = [0] * len(cabecalho)  # Inicializa uma lista para somar as colunas numéricas
        if coluna_numerica:
           # Identifica o índice da coluna numérica
           if coluna_numerica not in cabecalho:
              print(f"A coluna '{coluna_numerica}' não foi encontrada.")
              return
           indice_coluna_numerica = cabecalho.index(coluna_numerica)
        # Itera sobre cada linha (registro) e soma os valores das colunas numéricas
        for registro in registros:
            for i, valor in enumerate(registro):
                try:
                    if coluna_numerica == 0:
                       if isinstance(valor, (int, float)):
                          indice_coluna_numerica = i
                       soma_colunas[i] += float(valor)  # Soma os valores numéricos
                    else:
                        soma_colunas[indice_coluna_numerica] += float(registro[indice_coluna_numerica])  # Soma os valores numéricos
                except ValueError:
                    pass  # Se não for numérico, ignora a soma
        # O último registro com a quantidade e a soma das colunas
        ultimo_registro = []
        # Para cada coluna, coloca a soma na coluna numérica e vazio nas demais
        for i, coluna in enumerate(cabecalho):
            if i == indice_coluna_numerica:
                ultimo_registro.append(f"{soma_colunas[i]:>10.2f}" if isinstance(soma_colunas[i], (int, float)) else "")
            else:
                ultimo_registro.append("")  # Deixa vazio nas demais colunas
        # Adiciona o total de registros na primeira posição
        ultimo_registro[0] = f"{quantidade_registros} itens"
        registros.append(tuple(ultimo_registro))  # Adiciona o último registro na lista de registros
        # ==============================================================
        print(tabulate(registros, headers=cabecalho, tablefmt="grid"))
    elif isinstance(registros, list):
        if not cabecalho:
           cabecalho = []
           for index, item in enumerate(registros):
               coluna = f"coluna {index+1}"
               cabecalho.append(coluna)
        # =========================================================================
        # Inicializando a soma e contagem
        soma = 0
        quantidade_registros = len(registros)
        # Encontrar o primeiro campo numérico
        for registro in registros:
            for i, valor in enumerate(registro):
                if not coluna_numerica:
                   if isinstance(valor, (int, float)):  # Verificando se é numérico
                      soma += valor
                      break  # Só soma o primeiro valor numérico encontrado por registro
                else:
                    if cabecalho[i] == coluna_numerica:
                       if isinstance(valor, (int, float)):  # Verificando se é numérico
                          soma += valor
                       break  # Só soma o primeiro valor numérico encontrado por registro
            else:
                continue  # Se não encontrou número, continua para o próximo registro

        # Adiciona a última linha com a quantidade de registros e soma do primeiro campo numérico
        ultimo_registro = [""] * len(cabecalho)  # Cria uma lista de valores vazios
        ultimo_registro[0] = f"{quantidade_registros} itens"  # Coloca a quantidade na primeira coluna

        # Soma o primeiro campo numérico encontrado
        for i, registro in enumerate(registros):
            for i,valor in enumerate(registro):
                if not coluna_numerica:
                   if isinstance(valor, (int, float)):  # Soma o primeiro campo numérico
                      ultimo_registro[i] = f"{soma:>10.2f}"  # Coloca a soma na segunda coluna
                      break  # Se já encontrou e somou, sai do loop
                else:
                    if cabecalho[i] == coluna_numerica:
                       ultimo_registro[i] = f"{soma:>10.2f}"  # Coloca a soma na segunda coluna
                       break  # Se já encontrou e somou, sai do loop

        registros.append(ultimo_registro)  # Adiciona a última linha        
        # =========================================================================
        print(tabulate(registros, headers=cabecalho, tablefmt="grid"))
    elif isinstance(registros, pd.DataFrame):
        cabecalho = list(registros.columns)
        registros = registros.values.tolist()
        # =========================================================================
        # print(f"cabecalho: {cabecalho}")
        coluna_numerica_posicao = -1
        for i, coluna in enumerate(cabecalho):
            if coluna_numerica:
               if coluna.lower() == coluna_numerica.lower():
                  coluna_numerica_posicao = i 

        quantidade_registros = len(registros)
        soma = 0
        for col in registros:
            for i, valor in enumerate(col):
                if not coluna_numerica:
                   if isinstance(valor, (int, float)):
                      coluna_numerica_posicao = i
                      # registros[i].index(valor)
                      break
            if coluna_numerica is not None:
               break

        # Soma da coluna numérica
        for registro in registros:
            #print(f"registro: {registro}")
            for i, valor in enumerate(registro):
                if i == coluna_numerica_posicao:
                   soma += valor

        # Adiciona o último registro (quantidade e soma da primeira coluna numérica)
        ultimo_registro = [""] * len(cabecalho)
        ultimo_registro[0] = f"{quantidade_registros} itens"  # Coloca a quantidade de registros
        ultimo_registro[coluna_numerica_posicao] = f"{soma:>10.2f}"  # Soma da primeira coluna numérica

        registros.append(ultimo_registro)  # Adiciona o último registro        
        # =========================================================================
        print(tabulate(registros, headers=cabecalho, tablefmt="grid"))
    elif isinstance(registros, pd.Series):
        if not cabecalho:
            cabecalho = [registros.name if registros.name else "Valor"]
        registros = [[item] for item in registros.tolist()]
        # =========================================================================
        # print(f"cabecalho: {cabecalho}")
        coluna_numerica_posicao = -1
        for i, coluna in enumerate(cabecalho):
            if coluna_numerica:
               if coluna.lower() == coluna_numerica.lower():
                  coluna_numerica_posicao = i 
        quantidade_registros = len(registros)
        soma = 0
        for col in registros:
            for i, valor in enumerate(col):
                if not coluna_numerica:
                   if isinstance(valor, (int, float)):
                      coluna_numerica_posicao = i
                      # registros[i].index(valor)
                      break
            if coluna_numerica is not None:
               break

        # Soma da coluna numérica
        for registro in registros:
            for i, valor in enumerate(registro):
                if i == coluna_numerica_posicao:
                   soma += valor

        # Adiciona o último registro (quantidade e soma da primeira coluna numérica)
        ultimo_registro = [""] * len(cabecalho)
        ultimo_registro[0] = f"{quantidade_registros} itens"  # Coloca a quantidade de registros
        ultimo_registro[coluna_numerica_posicao] = f"{soma:>10.2f}"  # Soma da primeira coluna numérica
        registros.append(ultimo_registro)  # Adiciona o último registro        
        # =========================================================================
        print(tabulate(registros, headers=cabecalho, tablefmt="grid"))
    else:
        print(f"Formato de 'registros' {type(registros)} não suportado.")

def link_id(url):
    # Expressão regular para encontrar o ID
    url_id = re.search(r"(?:folders|file)/([a-zA-Z0-9_-]+)", url)
    id = ""
    if url_id:
        # Extrai o ID
        id = url_id.group(1)
        print(f"\nLink: {url}")
        print(f"   id: {id}")
    else:
        print("\nID não encontrado na URL:")
        print(f"{url}")
    return id

# Função para verificar se a pasta com o nome já existe em um determinado local
def pasta_existe(service_drive, pasta_nome, parent_id=None):
    # Pesquisa no Google Drive para encontrar pastas dentro de um parent específico (se fornecido)
    query = f"name = '{pasta_nome}' and mimeType = 'application/vnd.google-apps.folder'"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    results = service_drive.files().list(q=query).execute()
    if results.get('files', []):
        print(f"\nPasta '{pasta_nome}' já existe!")
        return results['files'][0]['id']  # Retorna o ID da pasta
    return None

# Função para verificar se a planilha com o nome já existe no Google Drive
def planilha_existe(service_drive, planilha_nome, parent_id):
    # Pesquisa no Google Drive para encontrar arquivos do tipo Google Sheets
    query = f"name = '{planilha_nome}' and mimeType = 'application/vnd.google-apps.spreadsheet'"
    if parent_id:
        query += f" and '{parent_id}' in parents"  # Adiciona o filtro de pasta pai se fornecido
    results = service_drive.files().list(q=query).execute()
    if results.get('files', []):
        print(f"\nPlanilha '{planilha_nome}' já existe!")
        return results['files'][0]['id']  # Retorna o ID da planilha
    return None

def criar_pasta(service_drive, pasta_nome, parent_id=None, ):
    # Verificar se a pasta já existe no local especificado
    pasta_id = pasta_existe(service_drive, pasta_nome, parent_id)
    if pasta_id:
        # Se a pasta já existe, apenas adicionar a permissão
        permissoes_pasta_arquivo(service_drive, pasta_id, 'anyone', 'reader', )
        return pasta_id

    # Se não existir, criar a pasta dentro do parent especificado (se houver)
    file_metadata = {
        'name': pasta_nome,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    if parent_id:
        file_metadata['parents'] = [parent_id]

    pasta = service_drive.files().create(body=file_metadata).execute()
    pasta_id = pasta['id']
    print(f"\nCriando pasta '{pasta_nome}'...")
    permissoes_pasta_arquivo(service_drive, pasta_id, 'anyone', 'writer')
    return pasta_id

# Atualizando a função de criar planilha para receber o nome da pasta (que agora será o parent_id)
def criar_planilha(service_drive, service_sheets, planilha_nome, parent_id=None):
    # Verificar se a planilha já existe
    planilha_id = planilha_existe(service_drive, planilha_nome, parent_id)
    if planilha_id:
        return planilha_id

    # Se não existir, criar a planilha
    body = {
        'properties': {'title': planilha_nome}
    }
    planilha = service_sheets.spreadsheets().create(body=body).execute()
    planilha_id = planilha['spreadsheetId']
    print(f"\nCriando planilha '{planilha_nome}'...")

    if parent_id:
        # Para definir a pasta pai de uma planilha, você precisa usar a API Drive
        service_drive.files().update(
            fileId=planilha_id,
            addParents=parent_id,
            fields='parents'  # Opcional: especifica os campos que você quer na resposta
        ).execute()

    return planilha_id

def permissoes_pasta_arquivo(service_drive, id, tipo, funcao):
    """
    Para type (tipo):
        • 'user': Acesso a um usuário específico por e-mail.
        • 'group': Acesso a um grupo do Google.
        • 'domain': Acesso a todos os usuários dentro de um domínio.
        • 'anyone': Acesso para qualquer pessoa com o link (sem conta Google).     
    Para role (função):
        • 'reader': Somente leitura.
        • 'commenter': Pode comentar, mas não editar.
        • 'writer': Pode editar e compartilhar.
        • 'owner': Propriedade total, com permissões para transferir.
    """
    # identificação do id (nome e tipo)
    informacoes = informacoes_driver(service_drive, id)
    # Verificar se é um arquivo ou uma pasta
    if informacoes[1] == 'application/vnd.google-apps.folder':
        tipo_item = 'Pasta'
    else:
        tipo_item = 'Arquivo'

    mensagem = f"{tipo_item}: {informacoes[0]} \nPermissões de '{tipo}' com função de '{funcao}'."

    # definindo as permissões
    permissao = {
        'type': tipo,
        'role': funcao
    }

    # Criando a permissão para a pasta
    try:
        service_drive.permissions().create(
            fileId=id,
            body=permissao
        ).execute()
        print(f"\n{mensagem}")
    except HttpError as error:
        print(f"\nPermissões NÂO gravadas '{mensagem}'")
        print(f"Falha: {error}")

def informacoes_driver(service_drive, id):
    # Buscando as informações do arquivo/pasta usando o ID
    conteudo = service_drive.files().get(fileId=id, fields='name,mimeType').execute()
    conteudo_nome = conteudo_tipo = ''
    if conteudo:
       conteudo_nome = conteudo['name']
       conteudo_tipo = conteudo['mimeType']     
    return [conteudo_nome, conteudo_tipo]

def planilha_dados(service_sheets, id, intervalo, cabecalho):
    try:
        # Inserir os cabeçalhos na planilha (a primeira linha)
        request = service_sheets.spreadsheets().values().update(
            spreadsheetId=id,
            range=intervalo,  # Intervalo onde os dados serão inseridos (exemplo: 'A1:C1')
            valueInputOption='RAW',  # Ou 'USER_ENTERED' se quiser que o Google Sheets interprete de forma inteligente
            body={'values': cabecalho}  # Dados a serem inseridos (os cabeçalhos)
        )
        response = request.execute()  # Executa a requisição para atualizar a planilha
        print(f"\nDados inseridos com sucesso no intervalo: {intervalo}")
        return response  # Retorna a resposta caso precise de mais informações (opcional)
    
    except Exception as e:
        print(f"\nErro ao inserir cabeçalhos: {e}")
        return None

def apagar_pasta_arquivo(service_drive, id=None, nome=None, parent_id=None):
    # Verificar se os parâmetros obrigatório foram passados
    if id == None and nome == None:
       print("A função 'apagar_pasta_arquivo' precisa de um 'ID' ou 'Nome' para continuar.")
       return # Adicionado return para sair da função

    # Verificar se o parâmetro fornecido é um ID ou um nome de pasta
    if nome:
        # Caso contrário, é um nome e precisamos buscar o ID da pasta
        id = pasta_existe(service_drive, nome, parent_id)
        if not id:
           id = planilha_existe(service_drive, nome, parent_id)
        if not id:
           print(f"\nNão foi encontrado 'pasta' ou 'planilha' com o nome '{nome}'.")
           return

    # Obter as informações da pasta
    informacoes = informacoes_driver(service_drive, id)
    
    if informacoes[1] == 'application/vnd.google-apps.folder':
        # Se for uma pasta, precisamos verificar se ela contém arquivos
        query = f"'{id}' in parents"
        results = service_drive.files().list(q=query).execute()
        
        if not results.get('files', []):
            # A pasta está vazia, podemos apagá-la diretamente
            try:
                service_drive.files().delete(fileId=id).execute()
                print(f"\nPasta '{informacoes[0]}' foi removida com sucesso!")
            except HttpError as error:
                print(f"\nErro ao remover a pasta '{informacoes[0]}': {error}")
        else:
            # A pasta contém arquivos, perguntar se deseja apagar os arquivos
            print(f"\nA pasta '{informacoes[0]}' contém arquivos. Você deseja apagar os arquivos dentro dela?")
            resposta = input("Digite 'sim' para apagar os arquivos e a pasta: ").strip().lower()
            
            if resposta == 'sim':
                # Apagar os arquivos dentro da pasta antes de apagar a pasta
                for arquivo in results['files']:
                    try:
                        service_drive.files().delete(fileId=arquivo['id']).execute()
                        print(f"   Arquivo '{arquivo['name']}' apagado com sucesso!")
                    except HttpError as error:
                        print(f"\nErro ao apagar o arquivo '{arquivo['name']}': {error}")
                
                # Agora apagar a pasta
                try:
                    service_drive.files().delete(fileId=id).execute()
                    print(f"\nPasta '{informacoes[0]}' foi removida com sucesso, juntamente com seus arquivos!")
                except HttpError as error:
                    print(f"\nErro ao remover a pasta '{informacoes[0]}': {error}")
            
            else:
                print("\nA exclusão da pasta foi cancelada.")
    else:
        # Se for um arquivo, podemos apagá-lo diretamente
        try:
            service_drive.files().delete(fileId=id).execute()
            print(f"\nArquivo '{informacoes[0]}' foi removido com sucesso!")
        except HttpError as error:
            print(f"\nErro ao remover o arquivo '{informacoes[0]}': {error}")

def intervalo_para_indices(intervalo_str):
    # O formato do intervalo é algo como 'B3:H9'
    
    # Primeiro, dividimos a string pelo ':'
    start_str, end_str = intervalo_str.split(':')
    
    # Depois, extraímos as colunas e as linhas separadamente
    start_col_str = start_str[0]  # Coluna inicial (ex: 'B')
    start_row_str = start_str[1:]  # Linha inicial (ex: '3')
    end_col_str = end_str[0]      # Coluna final (ex: 'H')
    end_row_str = end_str[1:]      # Linha final (ex: '9')

    # Converter as letras das colunas para números (A=0, B=1, ...)
    start_col = ord(start_col_str.upper()) - ord('A')
    end_col = ord(end_col_str.upper()) - ord('A')
    
    # Converter as linhas para inteiros (ajustando para 0-index)
    start_row = int(start_row_str) - 1  # Ajuste para 0-index
    end_row = int(end_row_str)          # Não precisa subtrair, pois é a última linha
    
    return (start_row, end_row, start_col, end_col)

def planilha_formatacao(service_sheets, planilha_id, aba_nome, cabecalho_intervalo, dados_intervalo):
    requests = []

    # 1. Remover as linhas de grade (gridlines)
    requests.append({
        'updateSheetProperties': {
            'properties': {
                'sheetId': 0,  # O ID da primeira aba (geralmente 0)
                'gridProperties': {
                    'hideGridlines': True
                }
            },
            'fields': 'gridProperties.hideGridlines'
        }
    })

    # 2. Intervalo do cabeçalho com fundo cinza
    start_row_cabecalho, end_row_cabecalho, start_col_cabecalho, end_col_cabecalho = intervalo_para_indices(cabecalho_intervalo)
    requests.append({
        'updateCells': {
            'range': {
                'sheetId': 0,
                'startRowIndex': start_row_cabecalho, 
                'endRowIndex': end_row_cabecalho,
                'startColumnIndex': start_col_cabecalho,
                'endColumnIndex': end_col_cabecalho + 1
            },
            'fields': 'userEnteredFormat.backgroundColor',
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredFormat': {
                                'backgroundColor': {
                                    'red': 0.85,  # Exemplo de cinza claro
                                    'green': 0.85,
                                    'blue': 0.85
                                }
                            }
                        }
                    ] * (end_col_cabecalho - start_col_cabecalho + 1)  # Ajusta o número de colunas
                }
            ]
        }
    })

    # 3. Adicionar bordas para o cabeçalho
    requests.append({
        'updateBorders': {
            'range': {
                'sheetId': 0,
                'startRowIndex': start_row_cabecalho,
                'endRowIndex': end_row_cabecalho,
                'startColumnIndex': start_col_cabecalho,
                'endColumnIndex': end_col_cabecalho+1
            },
            'top': {'style': 'SOLID', 'width': 1},
            'bottom': {'style': 'SOLID', 'width': 1},
            'left': {'style': 'SOLID', 'width': 1},
            'right': {'style': 'SOLID', 'width': 1},
            'innerHorizontal': {'style': 'SOLID', 'width': 1},
            'innerVertical': {'style': 'SOLID', 'width': 1}
        }
    })

    # 4. Adicionar bordas para os dados
    start_row_dados, end_row_dados, start_col_dados, end_col_dados = intervalo_para_indices(dados_intervalo)
    requests.append({
        'updateBorders': {
            'range': {
                'sheetId': 0,
                'startRowIndex': start_row_dados,
                'endRowIndex': end_row_dados,
                'startColumnIndex': start_col_dados,
                'endColumnIndex': end_col_dados+1
            },
            'top': {'style': 'SOLID', 'width': 1},
            'bottom': {'style': 'SOLID', 'width': 1},
            'left': {'style': 'SOLID', 'width': 1},
            'right': {'style': 'SOLID', 'width': 1},
            'innerHorizontal': {'style': 'SOLID', 'width': 1},
            'innerVertical': {'style': 'SOLID', 'width': 1}
        }
    })

    # 5. Alterar o nome da aba
    requests.append({
        'updateSheetProperties': {
            'properties': {
                'sheetId': 0, # O ID da primeira aba
                'title': aba_nome # Novo nome da aba
            },
            'fields': 'title'
        }
    })

    # 6. Centralizar os conteúdos do cabeçalho
    requests.append({
        'updateCells': {
            'range': {
                'sheetId': 0,
                'startRowIndex': start_row_cabecalho,
                'endRowIndex': end_row_cabecalho,
                'startColumnIndex': start_col_cabecalho,
                'endColumnIndex': end_col_cabecalho+1
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
                    ] * (end_col_cabecalho - start_col_cabecalho)
                }
            ]
        }
    })

    # 7. Reajustar o tamanho das colunas pelo tamanho dos cabeçalhos
    requests.append({
        'autoResizeDimensions': {
            'dimensions': {
                'sheetId': 0,
                'dimension': 'COLUMNS',            # Indica que queremos autoajustar colunas
                'startIndex': start_col_cabecalho, # Inicio
                'endIndex': end_col_cabecalho+1    # Final
            }
        }
    })
    
    try:
        body = {'requests': requests}
        response = service_sheets.spreadsheets().batchUpdate(
            spreadsheetId=planilha_id,
            body=body
        ).execute()
        print("\nFormatação aplicada com sucesso!")
        return response
    except HttpError as error:
        print(f"\nErro ao aplicar formatação: {error}")
        return None

# Função auxiliar para converter índice de coluna para letra de coluna do Excel
def coluna_numero_para_letra(numero):
    letra = ""
    while numero > 0:
        numero, restante = divmod(numero - 1, 26)
        letra = chr(65 + restante) + letra
    return letra

# Converter letra de coluna para número de coluna (baseado em 1)
def coluna_letra_para_numero(coluna_letra):
    numero = 0
    incremento = 1
    for char in reversed(coluna_letra):
        numero += (ord(char) - ord('A') + 1) * incremento
        incremento *= 26
    return numero

def planilha_celulas_intervalo(letra, numero, conteudo, tipo):
    if not conteudo:
       print("Função 'planilha_celulas_intervalo' parâmetro 'conteudo' incorreto.") 
       return None
    num_colunas = len(conteudo[0])
    num_linhas  = len(conteudo)+1
    coluna_inicial_numero = coluna_letra_para_numero(letra.upper())
    ultima_coluna_numero = coluna_inicial_numero + num_colunas - 1
    ultima_coluna_letra = coluna_numero_para_letra(ultima_coluna_numero)
    # Ajuste o range para o cabeçalho de forma dinâmica
    if tipo.upper() == 'C':
       #intervalo = f'{letra.upper()}{numero}:{letra.upper()}{ultima_coluna_numero}'
       intervalo = f'{letra.upper()}{numero}:{ultima_coluna_letra}{ultima_coluna_numero}'
    elif tipo.upper() == 'D':
       linha_final = int(numero) + num_linhas - 1
       numero = int(numero) + 1
       intervalo = f'{letra.upper()}{numero}:{ultima_coluna_letra}{linha_final}'
    else:
        print(f"Função 'planilha_celulas_intervalo' parâmetro 'tipo' inválido. Use 'C' para Cabeçalho ou 'D' para Dados.")
        return None
    return intervalo
