# WEB-SCRAPING-AUTOMATION

Carrying out web scraping on the Magazine Luiza website with the aim of automating processes and storing data in an Excel spreadsheet.

Web Scraping is the process of using computer programs to extract information from websites

# Objetivo: 

#### Extrair dados do site Magazine Luiza relacionados à pesquisa de notebooks, incluindo:

- Nome do produto
- Quantidade de avaliações
- URL do produto

# 🛠️ Bibliotecas Utilizadas: 

- selenium: Para automatizar a interação com o navegador e extrair dados da web.
- pandas: Para manipulação e análise de dados.
- re: Para realizar operações de expressão regular.
- openpyxl: Para manipulação de arquivos Excel.
- os: Para interagir com o sistema de arquivos.

### 🧩 Configuração do Navegador

```
webdriver.FirefoxOptions(): Configura o navegador Firefox para ser executado em modo headless (sem interface gráfica).

```
### 📊 Interação com o Excel 

#### - Conversão em DataFrames: Converte as listas melhores_data e piores_data em DataFrames do Pandas.

#### - Criação e Salvamento: Usa pd.ExcelWriter com o motor openpyxl para criar um arquivo Excel com duas abas:
 1. Melhores: Produtos com 100 ou mais avaliações.
 2. Piores: Produtos com menos de 100 avaliações.

#### - Ajuste de Dimensões: Ajusta as dimensões das colunas para uma melhor visualização usando a função adjust_column_widths(file_path).

## 🔍 Função para Extrair Dados da Página
### Processo Realizado pelo Código

1. Navegar até a URL fornecida.
2. Aguarda até que a lista de produtos esteja carregada.
3. Extrai dados de cada produto da pagina, incluindo o nome, quantidade de avaliações e URL.
4. Utiliza os dados coletados para fazer um comparativo das avaliações.
5. Adiciona produtos com 100 ou mais avaliações a uma lista de melhores , já os com menos de 100 avaliações vai para a lista de piores. 
6. Ajusta as dimensões do arquivo { Função adjust_column_widths(file_path) } e salva na pasta em formato Xlxs.

## 📚 Iteração sobre Páginas

O código itera sobre páginas numeradas de 1 a 17 para coletar dados e separa os produtos nas listas melhores_data e piores_data.
- Produtos sem avaliações são descartados. 

## 💻 Instruções de Uso



#### 1. Clone o meu Repositório:

#### 2. Instale as Dependências: 

```
pip install selenium pandas openpyxl
```
#### 3. Execute o Script:
```
python BootCode.py
```
O script coletará dados, filtrará os produtos e salvará as informações em dados.xlsx. 📁


