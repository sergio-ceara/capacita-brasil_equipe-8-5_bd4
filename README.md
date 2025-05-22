# Processador de Dados do Google Sheets

Este script Python se conecta ao Google Sheets para recuperar e processar dados das abas da planilha '[Banco 1]: Sensibilização e Prospecção', gerando um resumo de eventos e participantes relacionados à sensibilização, prospecção e qualificação. Os dados processados são então compilados em uma nova planilha do Google Sheets, que é criada e formatada dentro de uma pasta do projeto para ser utilizado no Dashboard no Looker Studio.

## Funcionalidades

* **Integração com Google Sheets**: Autentica e interage com o Google Sheets usando `gspread` e `google-api-python-client`.
* **Extração de Dados**: Lê dados de diferentes abas dentro de uma URL de planilha do Google Sheets especificada.
* **Transformação de Dados**:
    * Identifica e conta eventos por tipo (sensibilização, prospecção, qualificação) anualmente.
    * Conta indivíduos únicos ("Pessoas") participando de eventos de sensibilização, prospecção e qualificação anualmente.
    * Lida com conversões de data e remove entradas duplicadas para uma contagem precisa.
* **Integração com Google Drive**:
    * Cria uma nova subpasta em uma pasta compartilhada do Google Drive.
    * Gera uma nova planilha do Google Sheets dentro da subpasta criada.
    * Define as permissões apropriadas para a nova planilha.
    * Preenche a nova planilha com um cabeçalho estruturado e os dados processados.
    * Aplica formatação básica à planilha de saída.
* **Tratamento de Erros**: Inclui verificações básicas de conectividade com a internet e problemas com a autenticação e acesso à API do Google.

## Pré-requisitos

Antes de executar este script, certifique-se de ter o seguinte:

* **Python 3.x**: Instalado em seu sistema.
* **Projeto Google Cloud**: Um projeto Google Cloud com a API Google Sheets e a API Google Drive ativadas.
* **Conta de Serviço**: Uma conta de serviço com acesso às suas planilhas e ao Google Drive. Baixe o arquivo de chave JSON para esta conta de serviço.
* **Arquivo `.env`**: Um arquivo `.env` na raiz do diretório do seu projeto com as seguintes variáveis de ambiente configuradas:
    * `GOOGLE_CREDS_JSON_PATH`: O caminho para o arquivo de chave JSON da sua conta de serviço.
    * `BANCO_1_URL`: A URL da planilha do Google Sheets que contém seus dados brutos.
    * `PASTA_COMPARTILHADA`: A URL da pasta compartilhada do Google Drive onde a planilha de saída será criada.
    * `SUB_PASTA`: (Opcional) O nome da subpasta a ser criada dentro da pasta compartilhada.
    * `PLANILHA`: O nome desejado para a planilha de saída do Google Sheets.
    * `planilha_coluna_inicial`: (Opcional) A coluna inicial para inserção de dados (padrão: 'a').
    * `planilha_linha_inicial`: (Opcional) A linha inicial para inserção de dados (padrão: '1').

## Instalação

1.  **Clone o repositório (ou baixe o script):**

    ```bash
    git clone <url_do_repositorio>
    cd <diretorio_do_repositorio>
    ```

2.  **Instale as bibliotecas Python necessárias:**

    ```bash
    pip install gspread google-api-python-client pandas python-dotenv oauth2client
    ```

3.  **Configure seu arquivo `.env`:**

    Crie um arquivo chamado `.env` no mesmo diretório do seu script e adicione as variáveis de ambiente conforme descrito na seção **Pré-requisitos**.

## Uso

Execute o script a partir do seu terminal:

```bash
python nome_do_seu_script.py
```

O script imprimirá mensagens de status no console, incluindo a URL da nova planilha do Google Sheets assim que o processo for concluído.

## Estrutura do Script

* **Importações**: Bibliotecas essenciais para interação com a API do Google, manipulação de dados e carregamento de variáveis de ambiente.
* **`funcoes.py`**: (Assumido) Um módulo separado chamado `funcoes.py` contendo funções auxiliares para:
    * `verificar_conexao()`: Verifica a conectividade com a internet.
    * `banco_indisponivel()`: Lida com casos em que a planilha do Google Sheets está indisponível.
    * `listar_tabela()`: Formata e imprime tabelas de dados no console.
    * `driver_conexao()`: Estabelece conexões com os serviços do Google Drive e Sheets.
    * `link_id()`: Extrai o ID de uma URL do Google Drive.
    * `apagar_pasta_arquivo()`: Exclui pastas ou arquivos (usado para limpeza/teste no código fornecido).
    * `criar_pasta()`: Cria uma nova pasta no Google Drive.
    * `criar_planilha()`: Cria uma nova planilha do Google Sheets.
    * `permissoes_pasta_arquivo()`: Define permissões para arquivos/pastas.
    * `planilha_celulas_intervalo()`: Calcula os intervalos de células para solicitações da API Google Sheets.
    * `planilha_dados()`: Atualiza os dados em uma planilha do Google Sheets.
    * `planilha_formatacao()`: Aplica formatação a uma planilha do Google Sheets.
* **Bloco Principal de Execução**:
    * Verifica a conexão com a internet.
    * Carrega as variáveis de ambiente.
    * Autentica com as APIs do Google Sheets e Drive.
    * Abre a planilha de origem e itera através de suas abas.
    * Processa os dados das abas 'Dados de Inscrições em Eventos' e 'Dados de Prospecção e Qualificação' para contar eventos e pessoas.
    * Compila todos os dados agregados em um formato estruturado.
    * Chama funções auxiliares para criar uma nova pasta, uma nova planilha do Google Sheets, preenchê-la com dados e aplicar formatação.
    * Imprime a URL da planilha do Google Sheets gerada.
    * Inclui bloco `try-except` para exceções `HttpError` durante as interações com a API do Google.