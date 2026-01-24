import requests
import pandas as pd

base_url = "http://services.repcenter.skf.com"
port = ""  # A autenticação para esta API difere-se da API do MHV, tendo em vista que não utiliza API Key para autenticação e sim token de 20 min de duração
username = ""
password = ""
timeout = 60  # tempo de espera em segundos

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

# -------------------------------------
# DEFINIR FUNÇÃO PARA CHAMADA GET
# -------------------------------------
def get_data(endpoint):
    try:
        url = f"{base_url}:{port}{endpoint}"
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get(url, headers=headers, timeout=timeout)
        return r.json() if r.status_code == 200 else None
    except:
        return None

nextgen = get_data("/v1/nextgensensor")
df_nextgen = pd.DataFrame(nextgen) if nextgen else pd.DataFrame([{"erro": "Falha ao obter /v1/nextgensensor"}])
