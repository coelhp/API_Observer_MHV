import requests
import pandas as pd

base_url = "http://services.repcenter.skf.com"
port = ""
username = ""
password = ""
timeout = 60  # delay de espera em segundos

token_url = f"{base_url}:{port}/token"
payload = {
    "grant_type": "password",
    "username": username,
    "password": password
}
headers = {"Content-Type": "application/x-www-form-urlencoded"}

try:
    token_response = requests.post(token_url, data=payload, headers=headers, timeout=timeout)
    token = token_response.json().get("access_token") if token_response.status_code == 200 else None
except:
    token = None

#Função para obter os dados
def get_data(endpoint):
    try:
        url = f"{base_url}:{port}{endpoint}"
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get(url, headers=headers, timeout=timeout)
        return r.json() if r.status_code == 200 else None
    except:
        return None

gateways = get_data("/v1/gateways")
df_gateways = pd.DataFrame(gateways) if gateways else pd.DataFrame([{"erro": "Falha ao obter /v1/gateways"}])
