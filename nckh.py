from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import sys
import re
import time

sys.stdout.reconfigure(encoding='utf-8')

cOptions = Options()
cOptions.add_experimental_option("detach", True)


cService = ChromeService(executable_path="C:\\Program Files (x86)\\chromedriver.exe")
driver = webdriver.Chrome(service=cService, options=cOptions)


driver.get("https://www.lazada.vn/catalog/?spm=a211g0.flashsale.search.d_go&q=laptop")


try:
    products = []
    for i in range(102):
        content = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#root"))
        )

        soup = BeautifulSoup(driver.page_source, "html.parser")
        

        for item in soup.find_all('div', class_='Bm3ON'):  
            try:
                product_name = item.find('div', class_='RfADt').text.strip()
                product_price = item.find('span', class_='ooOxS').text.strip()
                product_sales_str = item.find('span', class_='_1cEkb').text.strip()
                match = re.search(r'\d+', product_sales_str)
                if match:
                    product_sales_num = int(match.group())
                else:
                    product_sales_num = 0
                
                products.append(
                    (product_name, product_price, product_sales_num)
                )    
            except AttributeError:  
                continue
        time.sleep(2)
        try:
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, "baxia-dialog-mask"))
            )
            next_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".ant-pagination-next > button"))
            )
            next_button.click()
        except Exception as e:
            print(f"Error during pagination: {e}")
        time.sleep(2)
        
    df = pd.DataFrame(products, columns=['Product name', 'Price', 'Sales'])
    print(df)
    
    df.to_excel('lazada_products.xlsx', index=False)  # Save as Excel file without the DataFrame index

finally:
    driver.quit()  # Close the browser properly
