import requests

cnpj = "27865757000102"  # Exemplo: Nubank
url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"

resp = requests.get(url)
print(resp.json())
