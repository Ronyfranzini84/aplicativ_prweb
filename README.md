# Aplicativo PRWeb

Aplicativo desktop em Python para automatizar transferencias de pedidos no PRWeb via Selenium.

O projeto possui dois modos de uso:
- Interface grafica em PySide6 (recomendado): `app_pyside6.py`
- Linha de comando (CLI): `cargas.py`

## Funcionalidades

- Importar planilha `.xlsx` com pedidos.
- Validar colunas obrigatorias (`Pedidos` e `desm`).
- Processar os pedidos no PRWeb com preenchimento automatico dos campos de transferencia.
- Registrar o retorno por linha na coluna `msg` (alertas, erros ou sucesso).
- Salvar automaticamente o resultado em arquivo `.xlsx`.
- Reabrir e analisar uma planilha de saida, com opcao de filtrar apenas linhas com mensagem.

## Estrutura do projeto

- `app_pyside6.py`: interface desktop, tabela de visualizacao e fluxo de importacao/processamento/analise.
- `cargas.py`: automacao Selenium, leitura/escrita de planilhas e modo CLI.
- `requirements.txt`: dependencias do projeto.
- `AppPRWeb.spec`: especificacao para build de executavel (PyInstaller).

## Requisitos

- Python 3.10+ (recomendado 3.11)
- Google Chrome instalado
- Acesso de rede ao PRWeb: `https://prweb01/bahia/gateway`

## Instalacao

1. Crie e ative um ambiente virtual.
2. Instale as dependencias:

```bash
pip install -r requirements.txt
```

## Uso - Interface grafica (recomendado)

Execute:

```bash
python app_pyside6.py
```

Fluxo na tela:

1. Clique em **Importar Entrada** e selecione a planilha `.xlsx`.
2. Defina o caminho de **Saida / Analise**.
3. Preencha os parametros de configuracao (empresa, login, senha, filial, atividade, motivo, carga).
4. Clique em **Processar no PRWeb**.
5. Acompanhe o status na barra inferior.
6. Ao final, o arquivo de saida sera salvo automaticamente.
7. Use **Analisar Saida** para reabrir o arquivo e, se quiser, marque **Exportar apenas linhas com mensagem**.

## Uso - Linha de comando (CLI)

Exemplo basico:

```bash
python cargas.py pedidos.xlsx --saida pedidos_resultado.xlsx
```

Exemplo com parametros:

```bash
python cargas.py pedidos.xlsx --empresa-gateway 29 --empresa-pr 21 --login 123456 --senha SUA_SENHA --filial 1200 --atividade D --motivo 35 --carga 17404047
```

Para ver todas as opcoes:

```bash
python cargas.py --help
```

## Formato da planilha de entrada

A planilha deve conter, no minimo, as colunas:

- `Pedidos`
- `desm`

Colunas adicionadas/atualizadas no resultado:

- `msg`: mensagem de alerta/erro/sucesso por linha.
- `status`: coluna de apoio (mantida pela normalizacao dos dados).

## Comportamento de salvamento

- O resultado e salvo em `.xlsx` usando `openpyxl`.
- Se o arquivo de saida estiver bloqueado (ex.: aberto no Excel), o sistema cria um arquivo alternativo com timestamp.

## Observacoes importantes

- O Chrome sera aberto durante a execucao da automacao.
- Os seletores Selenium dependem da tela atual do PRWeb; mudancas no sistema podem exigir ajuste de codigo.
- Evite versionar senhas reais em codigo, scripts ou arquivos compartilhados.
