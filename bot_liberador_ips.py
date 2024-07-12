from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd


# BOT PARA LIBERAÇÕES DE IPS

class BotLiberador:
    def __init__(self, username, password):
        # Corrigindo problema do Chrome fechando sozinho
        options = webdriver.ChromeOptions()
        options.add_experimental_option('detach', True)

        self.username = username
        self.password = password
        service=Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(options=options, service=service)

    # Entra no site do Sophos e loga
    def login(self):
        
        
        try:
            driver = self.driver
            driver.get('https://utm.geodados.com.br:4444/')
            
            time.sleep(1)
            
            # Autoriza o acesso do site
            driver.find_element(By.XPATH, '//button[@id="details-button"]').click()
            driver.find_element(By.XPATH, '//a[@id="proceed-link"]').click()
            
            time.sleep(3)
            
            # Loga no site            
            username_element = driver.find_element(By.XPATH, '//input[@name="login_username"]')
            #username_element.clear()
            username_element.send_keys(self.username)
            password_element = driver.find_element(By.XPATH, '//input[@name="login_password"]')
            #password_element.clear()
            password_element.send_keys(self.password)
            password_element.send_keys(Keys.RETURN)
        except Exception as e:
            print(f'Ocorreu um erro: {e}')
    
    
        # Entra no menu desejado
        time.sleep(20)
        
        # Definitions & Users
        driver.find_element(By.XPATH, '//div[@title="Definitions & Users"]').click()
        time.sleep(3)
        
        # Network Definitions
        driver.find_element(By.XPATH, '//div[@title="Network Definitions"]').click()
        time.sleep(4.5)
        
        
        # Insere os dados na tela
        df = pd.read_excel(caminho_planilha, sheet_name=planilha_selecionada)
        
        for i, linha in df.iterrows():
            # New Network Definition...
            driver.find_element(By.XPATH, '//div[@title="New Network Definition..."]').click()
            time.sleep(5)
            
            # Insere os valores dos nomes
            nome = driver.find_element(By.XPATH, '//input[@id="FORM_TABLE_definitions_networks_FORM_ELEMENT_name"]')
            nome.clear()
            nome.send_keys(linha['Name'])
            time.sleep(2)
            
            # Insere os valores dos IP's
            ip = driver.find_element(By.XPATH, '//input[@id="FORM_TABLE_definitions_networks_FORM_ELEMENT_address"]')
            ip.clear()
            ip.send_keys(linha['IPV4'])
            time.sleep(3)
            
            # Seleciona os valores das máscaras
            mascara_selecao = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//select[@id="FORM_TABLE_definitions_networks_FORM_ELEMENT_netmask"]'))
                )
            selecao = Select(mascara_selecao) # Cria um objeto Select
            option_selecao = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"//select[@id='FORM_TABLE_definitions_networks_FORM_ELEMENT_netmask']/option[@value='{linha['Máscara']}']"))
                ) # Espera até que a opção com o valor indicado no FOR esteja presente na página
            selecao.select_by_value(f"{linha['Máscara']}") # Seleciona a opção pelo valor indicado no FOR
            time.sleep(2.5)
            
            # Salva os dados inseridos
            driver.find_element(By.XPATH, '//div[@title="Save"]').click()
            time.sleep(5)
        
        
        driver.quit()
            

login = '**usuario**'
senha = '**senha**'
caminho_planilha = r'.\dados.xlsx'
planilha_selecionada = 'Planilha1'

iniciar = BotLiberador(login, senha)
iniciar.login()
