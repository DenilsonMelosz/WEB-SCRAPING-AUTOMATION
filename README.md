# WEB-SCRAPING-E-BOOT-
Carrying out web scraping on the Magazine Luiza website with the aim of automating processes and storing data in an Excel spreadsheet.

# Objetivo: 

### Extrair Dados do site Magazine Luiza , relacionados a pesquisa de notebooks

# Bibliotecas utilizadas: 

selenium: Para automatizar a interação com o navegador e extrair dados da web.
pandas: Para manipulação e análise de dados.
re: Para realizar operações de expressão regular.
openpyxl: Para manipulação de arquivos Excel.
os: Para interagir com o sistema de arquivos.

### Configuração do Navegador:

webdriver.FirefoxOptions(): Configura o navegador Firefox para ser executado em modo headless (sem interface gráfica).
Função para Extrair Dados da Página

### Interação com o Excel 
Pandas: Para manipulação e análise de dados. - Converte as listas de dados melhores_data e piores_data em DataFrames do Pandas.
Abre o arquivo Excel usando load_workbook.
Usa pd.ExcelWriter com o motor openpyxl para criar um arquivo Excel com duas abas:

### Processo realizado pelo codigo:
1º Navegar até a URL fornecida.
2º Aguarda até que a lista de produtos esteja carregada.
3º Extrai dados de cada produto da pagina, incluindo o nome, quantidade de avaliações e URL.
4º Utiliza os dados coletados para fazer um comparativo das avaliações.
5º Adiciona produtos com 100 ou mais avaliações a uma lista de melhores , já os com menos de 100 avaliações vai para a lista de piores. 
6º Ajusta as dimensões do arquivo { Função adjust_column_widths(file_path) } e salva na pasta em formato Xlxs.

Iteração sobre Páginas:
O código itera sobre páginas numeradas de 1 a 17, coletando dados de cada página.
Para cada página, o código chama extract_data_from_page(url) e separa os produtos em duas listas: melhores_data (com produtos bem avaliados) e piores_data (com produtos menos avaliados).
Salvar Dados em Arquivo Excel
Conversão em DataFrames:
