import requests

ddd = "45"  # Exemplo: Curitiba / PR
url = f"https://brasilapi.com.br/api/ddd/v1/{ddd}"

resp = requests.get(url)
print(resp.json())
