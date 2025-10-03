import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from credenciais import usuario, senha  # importa credenciais do arquivo separado

# =========================
# Configuração do navegador
# =========================
driver = webdriver.Chrome()
driver.maximize_window()

url = "https://epfa-dev1.fa.ocs.oraclecloud.com/"
print(f"Abrindo página: {url}")

inicio = time.time()
driver.get(url)
wait = WebDriverWait(driver, 30)

# =========================
# LOGIN
# =========================
try:
    # Espera pelo campo de usuário
    campo_usuario = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="idcs-signin-basic-signin-form-username"]')))
    campo_usuario.send_keys(usuario)

    campo_senha = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="idcs-signin-basic-signin-form-password|input"]')))
    campo_senha.send_keys(senha)

    botao_login = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idcs-signin-basic-signin-form-submit"]')))
    botao_login.click()

    fim = time.time()
    tempo_total = round(fim - inicio, 2)
    print(f"✅ Login realizado com sucesso em {tempo_total} segundos.")

except Exception as e:
    fim = time.time()
    tempo_total = round(fim - inicio, 2)
    print(f"❌ Erro durante o login após {tempo_total} segundos: {e}")
