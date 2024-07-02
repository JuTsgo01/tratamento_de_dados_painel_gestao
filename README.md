# Tratamento de Planilha para Painel de Gestão

Este projeto contém um script Python para tratar uma planilha Excel e gerar um arquivo CSV adequado para ser importado em uma base de dados. O script utiliza as bibliotecas `pandas` e `openpyxl` para manipulação de dados.

## Funcionalidades

- Carrega uma planilha Excel e processa os dados de acordo com regras específicas.
- Converte colunas específicas para tipos de dados apropriados (`float` e `int`).
- Substitui vírgulas por pontos em valores numéricos.
- Gera um arquivo CSV formatado com duas casas decimais para valores flutuantes.

## Objetivo

- Estamos com um projeto para enviar informações semanais ao banco de dados, e o arquivo deve ser enviado apenas em formato CSV.
- Outro ponto é que o arquivo passa por várias mãos; este script representa a última etapa antes do envio ao banco.
- Através do script, garantimos que as informações e o formato dos dados sejam imutáveis e sempre legíveis pelo banco.
- Ao garantir um fluxo dinâmico e imutável, asseguramos a integridade e consistência das informações ao longo de todo o processo.
  
## Pré-requisitos

- Python 3.6+
- Bibliotecas Python: `pandas`, `openpyxl`

## Instalação

1. Clone o repositório:
    ```sh
    git clone https://github.com/seu_usuario/nome_do_repositorio.git
    cd nome_do_repositorio
    ```

2. Crie um ambiente virtual (opcional, mas recomendado):
    ```sh
    python -m venv venv
    source venv/bin/activate  # No Windows use `venv\Scripts\activate`
    ```

3. Instale as dependências:
    ```sh
    pip install pandas openpyxl
    ```

## Uso

1. Coloque o arquivo Excel a ser processado no diretório do projeto.
2. Modifique o script `tratamentopainel.py` para definir o caminho correto do arquivo Excel.
3. Execute o script:
    ```sh
    python tratamentopainel.py
    ```
4. O arquivo CSV gerado será salvo no mesmo diretório.
