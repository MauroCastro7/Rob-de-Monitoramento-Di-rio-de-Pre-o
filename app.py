import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from datetime import datetime
import schedule

# Configuração do WebDriver
options = Options()
options.headless = True  # Executa o navegador em modo headless
driver_service = Service(ChromeDriverManager().install())

# URL do produto no site do Mercado Livre
product_url = "https://produto.mercadolivre.com.br/MLB-4859763490-xiaomi-poco-x6-pro-512-gb-12-ram-5g-nfc-nf-e-garantia-_JM"
# Nome do produto
product_name = "Xiaomi Poco X6 Pro"

# Nome do arquivo Excel
excel_file = "precos_produto.xlsx"

def fetch_price():
    driver = webdriver.Chrome(service=driver_service, options=options)
    driver.get(product_url)
    time.sleep(5)  # Espera carregar a página

    try:
        # Usando XPath para selecionar o elemento de preço
        price_element = driver.find_element(By.XPATH, '//*[@id="price"]/div/div[1]/div[1]/span/span/span[2]')
        
        if price_element:
            price_text = price_element.text
            price = float(price_text.replace('.', '').replace(',', '.').strip())
            driver.quit()
            return price
        else:
            print("Preço não encontrado")
            driver.quit()
            return None
    except Exception as e:
        print(f"Erro ao buscar o preço: {e}")
        driver.quit()
        return None

def update_excel():
    price = fetch_price()
    if price is not None:
        data = {
            'Produto': [product_name],
            'Data': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            'Valor': [price],
            'Link': [product_url]
        }

        df = pd.DataFrame(data)
        
        try:
            # Carregar o arquivo Excel existente
            book = load_workbook(excel_file)
            sheet = book.active
            
            # Adicionar novos dados ao final da planilha
            for index, row in df.iterrows():
                sheet.append(row.tolist())
            
            # Salvar as alterações no livro de trabalho
            book.save(excel_file)
        except FileNotFoundError:
            # Se o arquivo não existe, cria um novo
            df.to_excel(excel_file, index=False)

        print(f"Dados atualizados: {data}")
    else:
        print("Erro ao atualizar dados")

# Agendar a execução a cada 30 minutos
schedule.every(30).minutes.do(update_excel)

# Executar imediatamente e depois a cada 30 minutos
update_excel()

while True:
    schedule.run_pending()
    time.sleep(1)