import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from credenciais import usuario, senha  # importa credenciais do arquivo separado

# =========================
# Funções auxiliares
# =========================
def clicar_xpath(xpath, espera=1):
    """Clica em um elemento pelo XPATH e espera."""
    elemento = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    elemento.click()
    time.sleep(espera)

def digitar_xpath(xpath, texto, clear=False, espera=1):
    """Digita texto em um campo pelo XPATH e espera."""
    campo = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    if clear:
        campo.clear()
    campo.send_keys(texto)
    time.sleep(espera)

# =========================
# Planilha de despesas
# =========================
caminho_despesas = r"C:\ARQUIVOS TEMPORARIOS\RD11 Expences\Planilha\Automata.xlsx"
wb_desp = openpyxl.load_workbook(caminho_despesas, data_only=True)
ws_desp = wb_desp.active

# =========================
# Configuração do navegador
# =========================
driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://epfa-dev1.fa.ocs.oraclecloud.com/")
wait = WebDriverWait(driver, 30)  # espera de até 30s para elementos
time.sleep(60)

# =========================
# LOGIN
# =========================
driver.find_element(By.XPATH, '//*[@id="idcs-signin-basic-signin-form-username"]').send_keys(usuario)
time.sleep(60)
driver.find_element(By.XPATH, '//*[@id="idcs-signin-basic-signin-form-password|input"]').send_keys(senha)
time.sleep(60)
driver.find_element(By.XPATH, '//*[@id="idcs-signin-basic-signin-form-submit"]').click()
print("Login realizado, aguardando tela inicial...")
time.sleep(60)

# =========================
# Navegação até "Gerenciar Modelos de Relatório de Despesas"
# =========================
driver.find_element(By.XPATH, '//*[@id="pt1:_UIScmil2u"]').click()  # 1
time.sleep(60)

driver.find_element(By.XPATH, '//*[@id="pt1:_UIScmi4"]').click()     # 2
time.sleep(60)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]').click()  # 3
time.sleep(60)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]').send_keys("Gerenciar Modelos de Relatório de Despesas")  # 4
time.sleep(60)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:ctb3::icon"]').click()  # 5
time.sleep(60)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:AT1:_ATp:t1:1:cl4"]').click()  # 6
time.sleep(60)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1:value20::content"]').click()  # 7
time.sleep(60)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1:value20::content"]').send_keys("CORP_COM_BU")  # 8
time.sleep(60)

#9
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1::search"]').click() #Pesquisar
time.sleep(60)

#10 Administrativo
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:t1:1:commandLink1"]').click()
time.sleep(60) 