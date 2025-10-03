import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from credenciais import usuario, senha  # importa credenciais do arquivo separado

# Caminho da planilha de despesas
caminho_despesas = r"C:\ARQUIVOS TEMPORARIOS\RD11 Expences\Planilha\Automata.xlsx"
wb_desp = openpyxl.load_workbook(caminho_despesas, data_only=True)
ws_desp = wb_desp.active

# Inicia navegador
driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://epfa-dev1.fa.ocs.oraclecloud.com/fscmUI/faces/FuseTaskListManagerTop")
time.sleep(5)  # espera inicial maior

# === LOGIN ===
driver.find_element(By.XPATH, '//*[@id="userid"]').send_keys(usuario)
driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(senha)
driver.find_element(By.XPATH, '//*[@id="btnActive"]').click()
print("Login realizado, aguardando tela inicial...")
time.sleep(10)  # espera tela inicial carregar

# === NAVEGAÇÃO ATÉ GERENCIAR MODELOS DE RELATÓRIO ===
driver.find_element(By.XPATH, '//*[@id="pt1:_UIScmil2u"]').click()  # logo com iniciais do usuário
time.sleep(8)

driver.find_element(By.XPATH, '//*[@id="pt1:_UIScmi4"]').click()      # opção configuração e manutenção
time.sleep(8)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]').click()
time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]').send_keys("Gerenciar Modelos de Relatório de Despesas")
time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:ctb3::icon"]').click()  # botão pesquisar
time.sleep(8)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:AT1:_ATp:t1:1:cl4"]').click()
time.sleep(8)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1:value20::content"]').send_keys("CASS_AGRO_BU")
time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1::search"]').click()
time.sleep(8)

print("Todos os tipos de despesa foram cadastrados!")
