# Banco de Dados - [Banco 4] - Agregação de Valor

Este projeto automatiza a leitura, unificação, processamento e formatação de dados de duas abas distintas de uma planilha no Google Sheets. Ele cria uma nova planilha formatada e salva o resultado em uma pasta do Google Drive, com permissões públicas configuradas para ser usada num dashboard criado no Looker Studio.

---

## Funcionalidades

- Verificação de conectividade com a internet
- Conexão autenticada com Google Drive e Google Sheets via API
- Leitura segura de duas abas de uma planilha existente (Consultorias e Mentorias)
- Unificação e limpeza dos dados lidos
- Criação de nova planilha no Drive com:
  - Cabeçalhos destacados
  - Bordas e alinhamentos aplicados
  - Remoção de linhas de grade
  - Autoajuste de colunas
- Armazenamento em pasta compartilhada no Drive
- Permissões automáticas (anyone, edit)
- Variáveis de ambiente configuráveis via `.env`

---

## Estrutura do Projeto
├── main.py     # Arquivo principal de execução
├── funcoes.py  # Módulo com todas as funções auxiliares
├── .env        # Conteúdo particular (não subir para o GitHub)
└── README.md   # Este documento


---

## Requisitos

- Python 3.8+
- Conta Google com acesso à API do Drive e Sheets
- Credenciais do serviço (arquivo JSON da conta de serviço)
- Permissões de acesso à planilha e pasta no Drive

---

## Instalação

1. Clone o repositório:

```bash
git clone https://github.com/sergio-ceara/capacita-brasil_equipe-8-5_bd4.git
cd capacita-brasil_equipe-8-5_bd4
```

2. Instale as dependências:
```bash
pip install -r requirements.txt

arquivo: requirements.txt
  gspread==6.0.2
  google-api-python-client==2.126.0
  google-auth==2.29.0
  google-auth-oauthlib==1.2.0
  pandas==2.2.2
  python-dotenv==1.0.1
  openpyxl==3.1.2
```

3. Crie o arquivo .env com as seguintes variáveis:
```bash
GOOGLE_CREDS_JSON_PATH=/caminho/para/seu/arquivo.json
BANCO_4_URL=https://docs.google.com/spreadsheets/d/SEU_ID
PASTA_COMPARTILHADA=Nome da Pasta no Drive
SUB_PASTA=Nome da Subpasta (opcional)
PLANILHA=Nome da Nova Planilha a ser criada
```

4. Execução
```bash
python main.py
```

5. Exemplo de Uso
```bash
O script irá:
    Verificar a conexão com a internet
    Autenticar com Google APIs usando a conta de serviço
    Ler os dados de duas abas específicas da planilha
    Concatenar, limpar e formatar os dados
    Criar nova planilha com formatações visuais
    Exibir o link da nova planilha no terminal
```

6. Bibliotecas Utilizadas
```bash
    gspread
    google-api-python-client
    google-auth
    pandas
    python-dotenv
    openpyxl
```
## Junção com outros projetos
Este é 1 dos 5 módulos de leitura e transformação de dados que compõem um conjunto maior. Cada módulo trata de uma planilha diferente. A etapa final unificará todos os processamentos em uma execução única para geração de dashboards consolidados.
