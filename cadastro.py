import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException

def iniciar_driver():
    edge_options = Options()
    edge_options.use_chromium = True
    edge_options.add_argument('--ignore-certificate-errors')
    return webdriver.Edge(options=edge_options)

def acessar_cadastro(usr, pwd, cadastro):
    try:
        driver.get(f'http://{usr}:{pwd}@{cadastro}/')
        return True
    except WebDriverException:
        print(f"Erro ao acessar cadastro {cadastro}")
        return False

def check_checkbox(check_box):
    is_checked = check_box.get_attribute('checked')
    if not is_checked:  # Se não estiver marcado
        check_box.click()

def adicionar_cadastro(cadastro):
    print("Iniciando processo de adição de cadastros")
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'form')))
        tbody = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table[1]/tbody/tr[3]/td[3]/table/tbody/tr[2]/td/form/div[1]/div[2]')))
        
        lista_de_cadastro = ["", "", ""]

        for i in range(1, 4): 
            check_box = tbody.find_element(By.XPATH, f'./table[{i}]/tbody/tr[2]/td[2]/input')
            check_checkbox(check_box)

        for i, novo_cadastro in enumerate(lista_de_cadastro):
            input_element = tbody.find_element(By.XPATH, f'./table[{i+1}]/tbody/tr[1]/td[2]/input')

            input_element.clear()
            input_element.send_keys(novo_cadastro)
        
        aplicar()
    except TimeoutException:
        print("Não achou")
    except WebDriverException:
        print("Problema")

def aplicar():
    try: 
        driver.execute_script("""
            var div = document.getElementById('FORMULARIO');
            if (div) {
                div.style.zIndex = -1;
                div.style.position = 'absolute';
            }
        """)
        print("Ajustou o index")
        submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@value="Aplicar"]')))
        submit_button.click()
        atualizar_dados()
    except TimeoutException as e:
        print(f"Não encontrou o botão 'Aplicar': {e}")
    except WebDriverException as e:
        print(f"Erro ao clicar no botão 'Aplicar': {e}")

def atualizar_dados():
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[1]/tbody/tr[3]/td[1]/table/tbody/tr/td/a[9]'))).click()
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'Submit'))).click()
        # Atualizar a célula correspondente no DataFrame
        df_dados.at[index, 'STATUS'] = 'atualizado'
        df_dados.at[index, 'SISTEMA'] = 'CADASTRADO'
        df_dados.at[index, 'NOVO CADASTRO'] = ''
        df_dados.to_excel(file_path, sheet_name='DADOS', index=False, engine='openpyxl')
    except:
        print("Erro ao reiniciar")

# Carrega o arquivo Excel
file_path = ''
df_dados = pd.read_excel(file_path, sheet_name='DADOS', engine='openpyxl')

usr = ''
pwd = ""

driver = iniciar_driver()

for index, row in df_dados.iterrows():
    if row['Sistema'] != 'Não cadastrado' and row['STATUS'] != 'atualizado':    
        cadastro = row['cadastro']
        if acessar_cadastro(usr, pwd, cadastro):
            time.sleep(1)
            adicionar_cadastro(cadastro)

driver.quit()
