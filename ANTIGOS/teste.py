import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

url = "https://www.coagril-rs.com.br/"

resp = requests.get(url)
resp.encoding = "utf-8"
soup = BeautifulSoup(resp.text, "html.parser")

# Localiza a seção de cotações
cotacoes_div = soup.find("div", id="cotacao")
cards = cotacoes_div.find_all("div", class_="card_cotacao")

cotacoes = {}
for card in cards:
    nome = card.find("p", class_="nome_cotacao").get_text(strip=True)
    preco = card.find("p", class_="preco_cotacao").get_text(strip=True)
    cotacoes[nome] = preco

# DataFrame com timestamp
df = pd.DataFrame([cotacoes])
df["Data"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Salvar CSV acumulando histórico
arquivo = "cotacoes_coagril.csv"
try:
    df_old = pd.read_csv(arquivo)
    df_final = pd.concat([df_old, df], ignore_index=True)
except FileNotFoundError:
    df_final = df

df_final.to_csv(arquivo, index=False, encoding="utf-8-sig")

print("✅ Cotações coletadas com sucesso!")
print(df)
