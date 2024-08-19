from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re
from openpyxl import load_workbook
import os

# Configura o driver do navegador (usando o Firefox, como no seu exemplo)
options = webdriver.FirefoxOptions()
options.add_argument('--headless')  # Executa o navegador em modo headless (sem interface gráfica)
browser = webdriver.Firefox(options=options)

# Função para extrair dados de uma página específica
def extract_data_from_page(url):
    browser.get(url)
    data = []
    
    # Espera explícita para garantir que a lista de produtos esteja carregada
    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-testid='product-list']"))
    )
    
    product_list = browser.find_elements(By.CSS_SELECTOR, "div[data-testid='product-list'] li")
    for product in product_list:
        try:
            QTD_AVAL_TEXT = product.find_element(By.CSS_SELECTOR, "span.sc-kUdmhA.dWhxDa").text
            # Extrai o número de avaliações usando regex
            qtd_aval_match = re.search(r'\((\d+)\)', QTD_AVAL_TEXT)
            if qtd_aval_match:
                qtd_aval = int(qtd_aval_match.group(1))
                PRODUTO = product.find_element(By.CSS_SELECTOR, "a div h2").text
                URL = product.find_element(By.CSS_SELECTOR, "a").get_attribute('href')
                
                # Adiciona o item com base na quantidade de avaliações
                if qtd_aval >= 100:
                    data.append({'Nome': PRODUTO, 'QTD_AVAL': QTD_AVAL_TEXT, 'URL': URL})
                else:
                    continue
        
        except Exception as e:
            print(f"Erro ao processar o produto: {e}")
            continue
    
    return data

# Listas para armazenar os dados
melhores_data = []
piores_data = []

# Iterar sobre as páginas de 1 a 17
for page in range(1, 18):
    page_url = f'https://www.magazineluiza.com.br/busca/notebooks/?from=clickSuggestion&page={page}'
    print(f"Coletando dados da página {page}...")
    page_data = extract_data_from_page(page_url)
    for item in page_data:
        if int(re.search(r'\((\d+)\)', item['QTD_AVAL']).group(1)) >= 100:
            melhores_data.append(item)
        else:
            piores_data.append(item)

# Fecha o navegador
browser.quit()

# Convertendo as listas de dados em DataFrames do Pandas
df_melhores = pd.DataFrame(melhores_data)
df_piores = pd.DataFrame(piores_data)

# Renomear a coluna 'Nome' para 'PRODUTO'
df_melhores.rename(columns={'Nome': 'PRODUTO'}, inplace=True)
df_piores.rename(columns={'Nome': 'PRODUTO'}, inplace=True)

# Criar um escritor do Excel para salvar múltiplas abas
excel_path = 'dados.xlsx'
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    if not df_melhores.empty:
        df_melhores.to_excel(writer, sheet_name='Melhores', index=False)
    if not df_piores.empty:
        df_piores.to_excel(writer, sheet_name='Piores', index=False)

# Ajustar a largura das colunas
def adjust_column_widths(file_path):
    if not os.path.isfile(file_path):
        print(f"O arquivo {file_path} não existe.")
        return
    
    try:
        wb = load_workbook(filename=file_path)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            cell_length = len(str(cell.value))
                            if cell_length > max_length:
                                max_length = cell_length
                    except Exception as e:
                        print(f"Erro ao processar a célula: {e}")
                        continue
                adjusted_width = max_length + 2
                ws.column_dimensions[column_letter].width = adjusted_width
        wb.save(filename=file_path)
    except Exception as e:
        print(f"Erro ao ajustar a largura das colunas: {e}")

# Ajustar a largura das colunas no arquivo Excel
adjust_column_widths(excel_path)

print("Data received and saved to Excel with adjusted column widths")
