import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from credenciais import usuario, senha  # importa credenciais de arquivo separado

# ==============================
# FUNÇÕES AUXILIARES
# ==============================
def esperar(segundos=3):
    """Pausa o script para aguardar carregamento da tela."""
    time.sleep(segundos)


def clicar_xpath(driver, xpath, espera=3):
    """Localiza e clica em um elemento pelo XPath."""
    driver.find_element(By.XPATH, xpath).click()
    esperar(espera)


def digitar_xpath(driver, xpath, texto, espera=2):
    """Localiza e digita texto em um campo pelo XPath."""
    campo = driver.find_element(By.XPATH, xpath)
    campo.clear()
    campo.send_keys(texto)
    esperar(espera)


# ==============================
# CONFIGURAÇÕES
# ==============================
caminho_despesas = r"C:\ARQUIVOS TEMPORARIOS\RD11 Expences\Planilha\Automata.xlsx"
wb_desp = openpyxl.load_workbook(caminho_despesas, data_only=True)
ws_desp = wb_desp.active

# Inicia navegador
driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://epfa-dev1.fa.ocs.oraclecloud.com/fscmUI/faces/FuseTaskListManagerTop")
esperar(5)

# ==============================
# LOGIN
# ==============================
driver.find_element(By.XPATH, '//*[@id="userid"]').send_keys(usuario)
driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(senha)
clicar_xpath(driver, '//*[@id="btnActive"]', espera=10)
print("Login realizado!")

# ==============================
# NAVEGAÇÃO
# ==============================
clicar_xpath(driver, '//*[@id="pt1:_UIScmil2u"]', espera=8)   # logo usuário
clicar_xpath(driver, '//*[@id="pt1:_UIScmi4"]', espera=8)     # Configuração e manutenção

# Pesquisar funcionalidade
clicar_xpath(driver, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]', espera=2)
digitar_xpath(driver, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]',
              "Gerenciar Modelos de Relatório de Despesas")
clicar_xpath(driver, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:ctb3::icon"]', espera=8)

# Selecionar opção
clicar_xpath(driver, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:AT1:_ATp:t1:1:cl4"]', espera=8)

# Filtrar por BU
digitar_xpath(driver, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1:value20::content"]',
              "CASS_AGRO_BU")
clicar_xpath(driver, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1::search"]', espera=8)

# Selecionar modelo
clicar_xpath(driver, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:t1:11:commandLink1"]', espera=8)

print("Navegação concluída! Iniciando alterações...")

# ==============================
# ALTERAÇÃO DO MODELO
# ==============================

# 1 - Clicar na linha "Despesas De Viagem - Administrativo"
clicar_xpath(driver, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:t1::db"]/table/tbody/tr[2]/td[2]/div/table/tbody/tr/td[1]')

# 2 - Entrar na edição
clicar_xpath(driver, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:1:pt1:AP2:AT1:_ATp:table1:0:commandLink1"]')

# 3 - Selecionar "Obrigatório"
digitar_xpath(driver, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:2:pt1:AP2:socItemization::content"]',
              "Obrigatório")

# 4 - Marcar todos os checkboxes (0 até 15)
for i in range(16):
    xpath_checkbox = f'//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:2:pt1:AP2:AT3:_ATp:table3:{i}:chkItemize::Label0"]'
    try:
        clicar_xpath(driver, xpath_checkbox, espera=1)
    except Exception as e:
        print(f"Não encontrou checkbox {i}: {e}")

print("Processo finalizado: todos os checkboxes foram marcados!")
