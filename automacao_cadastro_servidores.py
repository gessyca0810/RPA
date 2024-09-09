import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException

# Função para iniciar o driver do Microsoft Edge com as opções configuradas
def iniciar_driver():
    options = Options()
    options.use_chromium = True
    options.add_argument('--ignore-certificate-errors')
    return webdriver.Edge(options=options)

# Função para acessar um cadastro específico usando autenticação básica
def acessar_cadastro(usr, pwd, cadastro):
    try:
        driver.get(f'http://{usr}:{pwd}@{cadastro}/')
        return True
    except WebDriverException:
        print(f"Erro ao acessar cadastro {cadastro}")
        return False

# Função para verificar e marcar uma checkbox se ela não estiver marcada
def marcar_checkbox(check_box):
    if not check_box.get_attribute('checked'):  # Se não estiver marcada
        check_box.click()

# Função para adicionar novos cadastros e marcar checkboxes
def adicionar_cadastro(cadastro):
    print("Iniciando processo de adição de cadastros")
    try:
        # Espera até que o formulário esteja pronto
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'form')))
        
        # Localiza a tabela contendo os campos de cadastro
        tbody = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/table[1]/tbody/tr[3]/td[3]/table/tbody/tr[2]/td/form/div[1]/div[2]')
        ))
        
        # Lista de novos cadastros
        lista_de_cadastro = ["", "", ""]

        # Marca as checkboxes para cada linha
        for i in range(1, 4): 
            check_box = tbody.find_element(By.XPATH, f'./table[{i}]/tbody/tr[2]/td[2]/input')
            marcar_checkbox(check_box)

        # Preenche os campos de texto com os novos cadastros
        for i, novo_cadastro in enumerate(lista_de_cadastro):
            input_element = tbody.find_element(By.XPATH, f'./table[{i+1}]/tbody/tr[1]/td[2]/input')
            input_element.clear()
            input_element.send_keys(novo_cadastro)
        
        aplicar_cadastros()

    except TimeoutException:
        print("Tempo esgotado ao tentar encontrar elementos")
    except WebDriverException:
        print("Problema ao interagir com a página")

# Função para aplicar os cadastros
def aplicar_cadastros():
    try: 
        # Ajusta o zIndex do formulário para permitir clicar no botão de aplicar
        driver.execute_script("""
            var div = document.getElementById('FORMULARIO');
            if (div) {
                div.style.zIndex = -1;
                div.style.position = 'absolute';
            }
        """)
        print("Index ajustado")

        # Clica no botão "Aplicar"
        aplicar_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@value="Aplicar"]')))
        aplicar_btn.click()

        # Atualiza os dados após aplicar os cadastros
        atualizar_dados()

    except TimeoutException as e:
        print(f"Erro: Não encontrou o botão 'Aplicar': {e}")
    except WebDriverException as e:
        print(f"Erro ao clicar no botão 'Aplicar': {e}")

# Função para atualizar o status no Excel
def atualizar_dados():
    try:
        # Clique nos botões para reiniciar e confirmar a ação
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[1]/tbody/tr[3]/td[1]/table/tbody/tr/td/a[9]'))).click()
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'Submit'))).click()

        # Atualiza as informações no DataFrame
        df_dados.at[index, 'STATUS'] = 'atualizado'
        df_dados.at[index, 'SISTEMA'] = 'CADASTRADO'
        df_dados.at[index, 'NOVO CADASTRO'] = 'Atualizado'
        df_dados.to_excel(file_path, sheet_name='DADOS', index=False, engine='openpyxl')

    except Exception as e:
        print(f"Erro ao atualizar os dados: {e}")

# Carrega o arquivo Excel com os dados dos cadastros
file_path = 'caminho_para_seu_arquivo.xlsx'
df_dados = pd.read_excel(file_path, sheet_name='DADOS', engine='openpyxl')

# Configurações de login
usr = ''
pwd = ''

# Inicia o driver do Edge
driver = iniciar_driver()

# Processa cada linha do DataFrame
for index, row in df_dados.iterrows():
    if row['Sistema'] != 'Não cadastrado' and row['STATUS'] != 'atualizado':
        cadastro = row['cadastro']
        if acessar_cadastro(usr, pwd, cadastro):
            time.sleep(1)
            adicionar_cadastro(cadastro)

# Encerra o driver
driver.quit()
