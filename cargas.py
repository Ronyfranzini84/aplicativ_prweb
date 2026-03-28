from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
import time  
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
driver.implicitly_wait(2)

driver.maximize_window()
driver.get("https://prweb01/bahia/gateway")

empresa = 29
emp = 21
login = 2634333
senha = "qwer7410"
filial = 1400
atividade = 'D'

def login_prweb(empresa, login, senha):
    campo_empresa = driver.find_element(By.XPATH, "/html/body/form[1]/table[1]/tbody/tr[1]/td[1]/input[1]")
    campo_empresa.send_keys(empresa)

    campo_login = driver.find_element(By.XPATH, "/html/body/form[1]/table[1]/tbody/tr[1]/td[2]/input[1]")
    campo_login.send_keys(login)

    campo_senha = driver.find_element(By.XPATH, "/html/body/form[1]/table[1]/tbody/tr[1]/td[3]/b[1]/input")
    campo_senha.send_keys(senha)

    botao_entrar = driver.find_element(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td[2]/input")
    botao_entrar.click()

# Realiza o login
login_prweb(empresa, login, senha)
time.sleep(5)

driver.find_element(By.XPATH, '/html/body/form[1]').click()

select_element = driver.find_element(By.NAME, "U01_DS_APL_VIS_SLC")
select = Select(select_element)

select.select_by_value("9")

driver.find_element(By.XPATH, '/html/body/form[1]/table[2]/tbody/tr/td[2]/input').click()
time.sleep(5)

def login_pr(emp, filial, atividade):
    campo_empresa = driver.find_element(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td[2]/input[1]")
    campo_empresa.clear()
    campo_empresa.send_keys(emp)

    campo_filial = driver.find_element(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td[3]/input[1]")
    campo_filial.send_keys(filial)

    campo_atividade = driver.find_element(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td[4]/input[1]")
    campo_atividade.send_keys(atividade)

login_pr(emp, filial, atividade)

driver.find_element(By.XPATH, '/html/body/form[1]/table[3]/tbody/tr/td[1]/table/tbody/tr[11]/td[1]/input').click()
driver.find_element(By.XPATH, '//*[@id="NM_BOT_CON"]').click() 
time.sleep(5)

driver.find_element(By.XPATH, '/html/body/form[1]/table[3]/tbody/tr[5]/td[1]/input').click()
driver.find_element(By.XPATH, '//*[@id="NM_BOT_PRC"]').click()
time.sleep(3)

driver.find_element(By.XPATH, '//*[@id="NM_BOT_TRA"]').click()
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="NM_BOT_TRA"]').click()
time.sleep(1)

driver.find_element(By.XPATH, '/html/body/form[1]/table[3]/tbody/tr/td[5]/input').click()
driver.find_element(By.XPATH, '/html/body/form[1]/table[3]/tbody/tr/td[5]/input').clear()
driver.find_element(By.XPATH, '/html/body/form[1]/table[3]/tbody/tr/td[5]/input').send_keys(empresa)
driver.find_element(By.XPATH, '/html/body/form[1]/table[3]/tbody/tr/td[6]/input').send_keys(login)

motivo = 35

# carregar excel
pedidos = pd.read_excel(r'C:\Users\2903471500\OneDrive - Grupo Casas Bahia S.A\Área de Trabalho\SLA\aplicativ_prweb\pedidos.xlsx')
# Adicionar coluna 'msg' se não existir
if 'msg' not in pedidos.columns:
    pedidos['msg'] = ''

carga = 16335512

for index, (pedido, desm) in enumerate(zip(pedidos['Pedidos'], pedidos['desm'])):
    # Se desm estiver vazio ou NaN, pula para a próxima linha
    if pd.isna(desm) or desm == "":
        continue

    # 1. campo senha
    driver.find_element(
        By.XPATH, 
        '/html/body/form[1]/table[3]/tbody/tr/td[7]/b/input'
    ).send_keys(senha)

    # 2. campo pedido
    driver.find_element(
        By.XPATH, 
        '/html/body/form[1]/table[4]/tbody/tr[1]/td[2]/input[1]'
    ).send_keys(pedido)

    # 3. campo desm
    driver.find_element(
        By.XPATH, 
        '/html/body/form[1]/table[4]/tbody/tr[1]/td[2]/input[2]'
    ).send_keys(desm)

    driver.find_element(By.XPATH, '//*[@id="NM_BOT_PRC"]').click()
    time.sleep(1)

    # 3. campo 46
    driver.find_element(
        By.XPATH, 
        '/html/body/form[1]/table[5]/tbody/tr[7]/td[2]/input[1]'
    ).send_keys(motivo)

    # 4. campo carga
    driver.find_element(
        By.XPATH, 
        '/html/body/form[1]/table[6]/tbody/tr/td/table/tbody/tr[1]/td[4]/input'
    ).send_keys(carga)

    time.sleep(1)

    # 5. botão processar
    driver.find_element(By.XPATH, '//*[@id="NM_BOT_PRC"]').click()
    time.sleep(1)

    # ENTER após processar
    try:
        alerta = WebDriverWait(driver, 10).until(EC.alert_is_present())
        texto_alerta = alerta.text
        alerta.accept()
        # Salvar mensagem na planilha
        pedidos.at[index, 'msg'] = texto_alerta

    except TimeoutException:
        print("Nenhum alerta apareceu.")
        # Salvar mensagem padrão na planilha
        pedidos.at[index, 'msg'] = "Nenhum alerta apareceu"

    time.sleep(1)

    # 6. botão limpar
    driver.find_element(By.XPATH, '//*[@id="NM_BOT_LIM"]').click()

    time.sleep(1)

print("Pedidos transferidos...")

# Salvar planilha atualizada com as mensagens de alerta
pedidos.to_excel(r'C:\Users\2903471500\OneDrive - Grupo Casas Bahia S.A\Área de Trabalho\SLA\aplicativ_prweb\pedidos_resultado.xlsx', index=False)
print("Resulados salvos em pedidos_resultado.xlsx")

driver.quit()