# Capacita Brasil - Indicadores do Banco 2

Este projeto automatiza o processo de leitura, tratamento e formatação de dados a partir da planilha '[Banco 2] - Seleção', com posterior criação de uma nova planilha formatada no Google Drive, para ser utilizada no dashboards.

## Funcionalidades

- Verificação de conexão com a internet.
- Conexão segura com APIs do Google Drive e Google Sheets.
- Leitura de dados da aba "Dados Seleção" da planilha '[Banco 2] - Seleção'.
- Transformação e padronização de dados:
  - Mapeamento de colunas.
  - Criação de coluna "País".
  - Conversão de status ("Sim"/"Não") em "Aprovado"/"Inscrito".
- Criação de nova planilha no Google Drive com dados tratados.
- Aplicação de formatações automáticas (cores, bordas, alinhamento, ajuste de colunas).
- Geração de link público para acesso à nova planilha.

## Estrutura do Projeto

```
capacita-brasil_banco-2_indicadores
│
├── capacita-brasil_banco-2_indicadores.py  # Script principal
├── funcoes.py                              # Módulo com funções auxiliares
├── .env                                    # Variáveis de ambiente (não incluso no repositório)
└── README.md                               # Esta documentação

````

## Pré-requisitos

- Python 3.8 ou superior
- Conta de serviço do Google com acesso à API do Drive e Sheets
- Credenciais JSON do Google (OAuth2)
- Planilha '[Banco 2] - Seleção' e aba `Dados Seleção`
- Biblioteca Python:
  - `gspread`
  - `pandas`
  - `openpyxl`
  - `google-api-python-client`
  - `oauth2client`
  - `python-dotenv`
  - `tabulate`

Instale com:

```bash
pip install -r requirements.txt
````

> Se ainda não houver um `requirements.txt`, crie com base nas dependências acima.

## Variáveis de Ambiente (.env)

Crie um arquivo `.env` com as seguintes variáveis:

```env
GOOGLE_CREDS_JSON_PATH=path/para/credenciais.json
BANCO_2_URL=https://docs.google.com/spreadsheets/d/EXEMPLO_DE_URL/edit
PASTA_COMPARTILHADA=https://drive.google.com/drive/folders/ID_DA_PASTA
SUB_PASTA=Indicadores
PLANILHA=Indicadores Banco 2
planilha_coluna_inicial=B
planilha_linha_inicial=2
```

## Como Executar

```bash
python capacita-brasil_banco-2_indicadores.py
```

A execução realizará os seguintes passos:

1. Checagem de conexão com a internet.
2. Leitura e validação da planilha de origem.
3. Criação de nova planilha no Google Drive.
4. Aplicação de formatação visual.
5. Impressão do link de acesso à nova planilha.

## Exemplo de Saída

```text
Conectado à internet.
Serviços do Google Drive, Sheets e cliente gspread ativos.
Nome da pasta compartilhada: ???
Nenhum registro repetido encontrado.
Criando planilha 'Indicadores Banco 2' com ID: 1A2B3C...
Formatação aplicada com sucesso!

Link da planilha criada: https://docs.google.com/spreadsheets/d/1A2B3C/edit
Fim.
```

## Possíveis Melhorias

* Utilizar este programa como referência para gerar outro com todos os indicadores.
