import openpyxl
import time
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# =========================
# Janela de Login (Tkinter)
# =========================
def abrir_janela_login():
    def confirmar():
        nonlocal_vars["usuario"] = entry_usuario.get()
        nonlocal_vars["senha"] = entry_senha.get()
        nonlocal_vars["bu"] = combo_bu.get()
        root.destroy()

    root = tk.Tk()
    root.title("Login Oracle Fusion")
    root.geometry("350x200")
    root.resizable(False, False)

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack(expand=True, fill="both")

    # Usuário
    tk.Label(frame, text="Usuário:").grid(row=0, column=0, sticky="w")
    entry_usuario = tk.Entry(frame, width=30)
    entry_usuario.grid(row=0, column=1)

    # Senha
    tk.Label(frame, text="Senha:").grid(row=1, column=0, sticky="w")
    entry_senha = tk.Entry(frame, width=30, show="*")
    entry_senha.grid(row=1, column=1)

    # BU
    tk.Label(frame, text="BU:").grid(row=2, column=0, sticky="w")
    combo_bu = ttk.Combobox(frame, width=27, values=[
        "PLU_AGRO_BU",
        "PLUMA_GENETICS_BU",
        "PLUSVAL_AGRO_BU",
        "DF_AVES_BU",
    ])
    combo_bu.grid(row=2, column=1)
    combo_bu.set("PLU_AGRO_BU")  # valor padrão

    # Botão
    btn = tk.Button(frame, text="Entrar", command=confirmar, bg="#4CAF50", fg="white")
    btn.grid(row=3, column=0, columnspan=2, pady=10)

    root.mainloop()
    return nonlocal_vars["usuario"], nonlocal_vars["senha"], nonlocal_vars["bu"]

nonlocal_vars = {"usuario": "", "senha": "", "bu": ""}
usuario, senha, bu_escolhida = abrir_janela_login()

# =========================
# Planilha
# =========================
caminho_despesas = r"C:\ARQUIVOS TEMPORARIOS\RD11 Expences\Planilha\Automata.xlsx"
wb_desp = openpyxl.load_workbook(caminho_despesas, data_only=True)
ws_desp = wb_desp.active

# =========================
# Selenium
# =========================
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 20)

driver.get("https://epfa-dev1.fa.ocs.oraclecloud.com/fscmUI/faces/FuseTaskListManagerTop")
time.sleep(5)

# =========================
# Funções auxiliares
# =========================
def clicar_xpath(xpath, espera=1, timeout=20):
    """Tenta clicar normalmente. Se falhar, força o clique via JavaScript."""
    try:
        el = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        try:
            el.click()
        except Exception:
            driver.execute_script("arguments[0].click();", el)
        time.sleep(espera)
    except TimeoutException as e:
        raise Exception(f"Elemento não encontrado/clicável: {xpath}") from e

def digitar_xpath(xpath, texto, clear=True, espera=0.5, timeout=20):
    try:
        el = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        if clear:
            try:
                el.clear()
            except Exception:
                pass
        el.send_keys(texto)
        time.sleep(espera)
    except TimeoutException as e:
        raise Exception(f"Campo não encontrado para digitar: {xpath}") from e

# === LOGIN ===
digitar_xpath('//*[@id="userid"]', usuario, clear=True, espera=0.2)
digitar_xpath('//*[@id="password"]', senha, clear=True, espera=0.2)
clicar_xpath('//*[@id="btnActive"]', espera=8)
print("Login realizado, aguardando tela inicial...")
time.sleep(10)

# === NAVEGAÇÃO ===
clicar_xpath('//*[@id="pt1:_UIScmil2u"]', espera=8)
clicar_xpath('//*[@id="pt1:_UIScmi4"]', espera=8)

clicar_xpath('//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]', espera=1)
digitar_xpath('//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:it2::content"]',
              "Gerenciar Modelos de Relatório de Despesas", clear=True, espera=0.5)
clicar_xpath('//*[@id="pt1:r1:0:r0:0:r1:0:AP1:s92:ctb3::icon"]', espera=8)

clicar_xpath('//*[@id="pt1:r1:0:r0:0:r1:0:AP1:AT1:_ATp:t1:1:cl4"]', espera=8)

digitar_xpath('//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1:value20::content"]',
              bu_escolhida, clear=True, espera=0.5)
clicar_xpath('//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:q1::search"]', espera=15)

clicar_xpath('//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:0:pt1:AP1:AT1:_ATp:t1::db"]/table/tbody/tr[3]/td[2]/div/table/tbody/tr/td[1]', espera=8)

# === REPETIR CRIAR TIPO DE DESPESA PARA LINHAS 1 A 20 ===
for linha in range(1, 21):
    try:
        valor_c = str(ws_desp[f"C{linha}"].value or "")
        valor_d = str(ws_desp[f"D{linha}"].value or "")
        valor_i = str(ws_desp[f"I{linha}"].value or "")

        print(f"Iniciando cadastro linha {linha} -> "
              f"C='{valor_c}' D='{valor_d}' I='{valor_i}'")

        # Clicar em Criar Tipo de Despesa
        clicar_xpath(
            '//*[@id="pt1:r1:0:rt:1:r2:0:dynamicRegion1:1:pt1:AP2:AT1:_ATp:cmdCreate::icon"]',
            espera=8
        )

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

        # Mensagem no console
        print(f"Linha {linha} cadastrada com sucesso!")

        # Mostra alerta visual no navegador
        driver.execute_script(f"alert('Linha {linha} cadastrada com sucesso!');")
        time.sleep(2)
        driver.switch_to.alert.accept()

        time.sleep(1)

    except Exception as e:
        print(f"Erro ao processar linha {linha}: {e}")
        continue

print("Todos os tipos de despesa processados (veja mensagens acima).")
