import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from credenciais import usuario, senha  # importa credenciais do arquivo separado

# Caminho da planilha de despesas
caminho_despesas = r"C:\ARQUIVOS TEMPORARIOS\RD11 Expences\Planilha\Automata.xlsx"
wb_desp = openpyxl.load_workbook(caminho_despesas, data_only=True)
ws_desp = wb_desp.active

# Inicia navegador Edge
driver = webdriver.Edge()
driver.maximize_window()
driver.get("https://epfa-dev1.fa.ocs.oraclecloud.com/fscmUI/faces/FuseTaskListManagerTop")
time.sleep(5)  # espera inicial maior

wait = WebDriverWait(driver, 30)  # espera explícita padrão

# === LOGIN ===
driver.find_element(By.XPATH, '//*[@id="userid"]').send_keys(usuario)
driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(senha)
driver.find_element(By.XPATH, '//*[@id="btnActive"]').click()
print("Login realizado, aguardando tela inicial...")
time.sleep(10)  # espera tela inicial carregar

# === NAVEGAÇÃO ATÉ GERENCIAR MODELOS DE RELATÓRIO ===
driver.find_element(By.XPATH, '//*[@id="pt1:_UIScmil2u"]').click()  # logo com iniciais do usuário
time.sleep(8)

driver.find_element(By.XPATH, '//*[@id="pt1:_UIScmi4"]').click()  # opção configuração e manutenção
time.sleep(8)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]').click()
time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]').send_keys(
    "Gerenciar Modelos de Relatório de Despesas")
time.sleep(2)

driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:ctb3::icon"]').click()  # botão pesquisar
time.sleep(8)

# === 3 OPÇÕES PARA CLICAR NO LINK ===
clicado = False

# OPÇÃO 1: XPath original
try:
    elemento = wait.until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:t1:2:commandLink1"]'
    )))
    elemento.click()
    clicado = True
    print("Clique realizado com XPath original (Opção 1).")
except Exception as e:
    print(f"Falhou na OPÇÃO 1: {e}")

# OPÇÃO 2: Buscar pelo texto visível
if not clicado:
    try:
        elemento = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//a[contains(text(),'Gerenciar Modelos de Relatório de Despesas')]"
        )))
        elemento.click()
        clicado = True
        print("Clique realizado pelo texto (Opção 2).")
    except Exception as e:
        print(f"Falhou na OPÇÃO 2: {e}")

# OPÇÃO 3: Dentro de iframe
if not clicado:
    try:
        iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)
        elemento = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//a[contains(text(),'Gerenciar Modelos de Relatório de Despesas')]"
        )))
        elemento.click()
        driver.switch_to.default_content()  # volta para o contexto principal
        clicado = True
        print("Clique realizado dentro do iframe (Opção 3).")
    except Exception as e:
        print(f"Falhou na OPÇÃO 3: {e}")

if not clicado:
    print("❌ Nenhuma das 3 opções conseguiu clicar no link. Verifique se o XPath/texto mudou.")

# === Funções auxiliares ===
def clicar_xpath(xpath, espera=0):
    el = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    el.click()
    if espera > 0:
        time.sleep(espera)

def digitar_xpath(xpath, valor, clear=True, espera=0):
    el = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    if clear:
        try:
            el.clear()
        except Exception:
            pass
    el.send_keys(valor)
    if espera > 0:
        time.sleep(espera)

# === REPETIR CRIAR TIPO DE DESPESA PARA LINHAS 1 A 20 ===
for linha in range(1, 21):
    try:
        valor_c = str(ws_desp[f"C{linha}"].value or "")
        valor_d = str(ws_desp[f"D{linha}"].value or "")
        valor_i = str(ws_desp[f"I{linha}"].value or "")

        print(f"Iniciando cadastro linha {linha} -> "
              f"C='{valor_c}' D='{valor_d}' I='{valor_i}'")

        # Clicar em Criar Tipo de Despesa
        clicar_xpath('//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:1:pt1:AP2:AT1:_ATp:cmdCreate::icon"]',
                     espera=8)

        # Preencher campos
        digitar_xpath('//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:2:pt1:AP2:inputText5::content"]',
                      valor_c, clear=True, espera=0.5)

        digitar_xpath('//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:2:pt1:AP2:inputText6::content"]',
                      valor_c, clear=True, espera=0.5)

        digitar_xpath('//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:2:pt1:AP2:socCategCode::content"]',
                      valor_d, clear=True, espera=0.5)

        # Campo de data
        campo_data = wait.until(EC.presence_of_element_located((By.XPATH,
            '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:2:pt1:AP2:TypeStartDate::content"]')))
        try:
            campo_data.clear()
        except Exception:
            pass
        campo_data.send_keys("01/01/1951")
        time.sleep(1)

        # Conta contábil
        digitar_xpath('//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:2:pt1:AP2:kfp1_AccountKff1IteratorcontaContabil23001::content"]',
                      valor_i, clear=True, espera=0.5)

        # Salvar e fechar
        clicar_xpath('//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:2:pt1:AP2:cmdSaveClose2"]/table/tbody/tr/td[1]/a',
                     espera=8)

        print(f"Linha {linha} cadastrada com sucesso!")
        driver.execute_script(f"alert('Linha {linha} cadastrada com sucesso!');")
        time.sleep(2)
        driver.switch_to.alert.accept()

        time.sleep(1)

    except Exception as e:
        print(f"Erro ao processar linha {linha}: {e}")
        continue

print("Todos os tipos de despesa processados (veja mensagens acima).")
