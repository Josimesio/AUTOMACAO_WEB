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

#1
driver.find_element(By.XPATH, '//*[@id="pt1:_UIScmil2u"]').click()  # logo com iniciais do usuário
time.sleep(10)

#2
driver.find_element(By.XPATH, '//*[@id="pt1:_UIScmi4"]').click()      # opção configuração e manutenção
time.sleep(10)

#3
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]').click()
time.sleep(10)

#4
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]').send_keys("Gerenciar Modelos de Relatório de Despesas")
time.sleep(10)

#5
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:ctb3::icon"]').click()  # botão pesquisar
time.sleep(10)

#6
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:r0:0:r1:0:AP1:AT1:_ATp:t1:1:cl4"]').click()
time.sleep(10)

#7
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1:value20::content"]').click()
time.sleep(10)

#8
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1:value20::content"]').send_keys("PLU_AGRO_BU")
time.sleep(10)

#9
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1::search"]').click() #Pesquisar
time.sleep(10)

#10
driver.find_element(By.XPATH, '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:t1:2:commandLink1"]').click()
time.sleep(10)

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

        # ⬅️ Intervalo de 8 segundos entre cada linha
        time.sleep(8)

    except Exception as e:
        print(f"Erro ao processar linha {linha}: {e}")
        continue


print("Todos os tipos de despesa processados (veja mensagens acima).")
