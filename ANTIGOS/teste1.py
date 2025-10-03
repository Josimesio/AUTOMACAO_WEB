import openpyxl
import subprocess
import pyautogui
import time


pyautogui.click(x=1757,y=412)
time.sleep(2)
pyautogui.click(x=1480,y=260)
time.sleep(2)
caminho = r"C:\ARQUIVOS TEMPORARIOS\RD11 Expences\Planilha\TEMPLATE_DESPESAS_V1.xlsx"

# Abre a planilha
wb = openpyxl.load_workbook(caminho, data_only=True)
ws = wb.active  # primeira aba
