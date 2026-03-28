Aplicativo PRWeb

Este projeto automatiza o processamento de pedidos no PRWeb com Selenium e agora possui uma interface em PySide6 para uso como aplicativo desktop.

Funcionalidades
- Importar planilha Excel com as colunas obrigatorias: Pedidos e desm.
- Executar o processamento no PRWeb usando os mesmos XPaths do script original.
- Registrar em msg o retorno de alerta (ou erro) por linha.
- Exportar resultado completo ou apenas linhas com mensagem.

Arquivos principais
- cargas.py: logica de automacao Selenium (refatorada para ser reutilizavel).
- app_pyside6.py: interface desktop com importacao/processamento/exportacao.

Instalacao
1. Criar/ativar ambiente virtual Python.
2. Instalar dependencias:

	pip install -r requirements.txt

Como executar

Modo aplicativo (recomendado)
1. Execute:

	python app_pyside6.py

2. Na tela:
- Selecione o arquivo Excel de entrada.
- Ajuste os campos de configuracao (empresa, login, senha, filial, etc).
- Clique em Processar no PRWeb.
- Exporte o resultado para um arquivo Excel.

Modo script legado
1. Coloque o arquivo pedidos.xlsx na pasta do projeto.
2. Execute:

	python cargas.py

3. O arquivo pedidos_resultado.xlsx sera gerado na pasta do projeto.

Observacoes
- O navegador Chrome sera aberto durante a execucao.
- A automacao depende do acesso interno ao endereco: https://prweb01/bahia/gateway.
- A coluna msg e usada para armazenar alerta/erro por linha processada.
